using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleDiff
{
    public partial class MainForm : Form
    {
        private readonly string _settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleDiff.json");
        private readonly BindingList<DiffResult> _diffResults = new BindingList<DiffResult>();
        private readonly string _logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleDiff.log");
        private readonly object _logLock = new object();
        private CancellationTokenSource _cancellationTokenSource;
        private int _lastLoggedProgressPercent = -1;
        private int _lastUiProgressPercent = -1;
        private volatile bool _isLogEnabled = true;
        private bool _isWinMergeComparisonRunning;
        private bool _isRunning;

        public MainForm()
        {
            InitializeComponent();
            InitializeCustomComponents();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
            btnStart.Click += btnStart_Click;
            btnStop.Click += btnStop_Click;
            btnRefDirSrc.Click += btnRefDirSrc_Click;
            btnRefDirDst.Click += btnRefDirDst_Click;
            btnRefWinMerge.Click += btnRefWinMerge_Click;
            btnRunWinMerge.Click += btnRunWinMerge_Click;
            dataGridView.CellDoubleClick += dataGridView_CellDoubleClick;
            dataGridView.CellToolTipTextNeeded += dataGridView_CellToolTipTextNeeded;
            dataGridView.KeyDown += dataGridView_KeyDown;
        }

        private void InitializeCustomComponents()
        {
            dataGridView.AutoGenerateColumns = false;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.MultiSelect = true;
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersVisible = false;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.ShowCellToolTips = true;
            clmDirPathSrc.DataPropertyName = nameof(DiffResult.SourceDirectory);
            clmFileNameSrc.DataPropertyName = nameof(DiffResult.SourceFileName);
            clmDirPathDst.DataPropertyName = nameof(DiffResult.DestinationDirectory);
            clmFileNameDst.DataPropertyName = nameof(DiffResult.DestinationFileName);
            dataGridView.DataSource = _diffResults;
            chkEnableSubDir.Checked = true;
            chkUseDebugLog.Checked = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;
            lblInfo.Text = "0/0ファイル";
            btnStop.Enabled = false;
            EnablePathDragDrop(txtWinMergePath, File.Exists);
            EnablePathDragDrop(txtDirPathSrc, Directory.Exists);
            EnablePathDragDrop(txtDirPathDst, Directory.Exists);
            chkUseDebugLog.CheckedChanged += chkUseDebugLog_CheckedChanged;
            UpdateLogEnabledState();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isRunning)
            {
                _cancellationTokenSource?.Cancel();
            }
            SaveSettings();
        }

        private void btnRefWinMerge_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "WinMerge|WinMergeU.exe|実行ファイル|*.exe|すべてのファイル|*.*";
                dialog.CheckFileExists = true;
                dialog.Title = "WinMergeの実行ファイルを選択してください";
                if (File.Exists(txtWinMergePath.Text))
                {
                    dialog.InitialDirectory = Path.GetDirectoryName(txtWinMergePath.Text);
                    dialog.FileName = Path.GetFileName(txtWinMergePath.Text);
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtWinMergePath.Text = dialog.FileName;
                }
            }
        }

        private void btnRefDirSrc_Click(object sender, EventArgs e)
        {
            SelectFolder(txtDirPathSrc);
        }

        private void btnRefDirDst_Click(object sender, EventArgs e)
        {
            SelectFolder(txtDirPathDst);
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning)
            {
                return;
            }

            var sourceRoot = txtDirPathSrc.Text.Trim();
            var destinationRoot = txtDirPathDst.Text.Trim();
            LogInfo($"開始ボタン押下 Source='{sourceRoot}' Destination='{destinationRoot}' IncludeSubDir={chkEnableSubDir.Checked}");

            if (!ValidateDirectory(sourceRoot, "比較元フォルダパス"))
            {
                LogInfo("比較元フォルダパスの検証に失敗し、処理を中断します。");
                return;
            }

            if (!ValidateDirectory(destinationRoot, "比較先フォルダパス"))
            {
                LogInfo("比較先フォルダパスの検証に失敗し、処理を中断します。");
                return;
            }

            try
            {
                sourceRoot = Path.GetFullPath(sourceRoot);
                destinationRoot = Path.GetFullPath(destinationRoot);
                LogInfo($"フルパス解決 Source='{sourceRoot}' Destination='{destinationRoot}'");
            }
            catch (Exception ex)
            {
                LogError("フォルダパスの解決に失敗しました。", ex);
                MessageBox.Show($"フォルダパスの解決に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var searchOption = chkEnableSubDir.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            List<FileItem> sourceFiles;
            List<FileItem> destinationFiles;

            try
            {
                LogInfo($"フォルダ走査開始 Source SearchOption={searchOption}");
                sourceFiles = EnumerateFileItems(sourceRoot, searchOption);
                LogInfo($"フォルダ走査開始 Destination SearchOption={searchOption}");
                destinationFiles = EnumerateFileItems(destinationRoot, searchOption);
                LogInfo($"フォルダ走査完了 SourceCount={sourceFiles.Count} DestinationCount={destinationFiles.Count}");
            }
            catch (Exception ex)
            {
                LogError("フォルダの走査に失敗しました。", ex);
                MessageBox.Show($"フォルダの走査に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (sourceFiles.Count == 0)
            {
                LogInfo("比較元フォルダにファイルが存在しません。");
                MessageBox.Show("比較元フォルダにファイルが存在しません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _diffResults.Clear();
            var total = sourceFiles.Count;
            progressBar.Maximum = total;
            progressBar.Value = 0;
            UpdateProgressLabel(0, total);
            _lastLoggedProgressPercent = -1;
            _lastUiProgressPercent = -1;
            LogProgressIfNeeded(0, total);
            LogInfo($"比較準備完了 TotalFiles={total}");

            var destinationMap = BuildFileMap(destinationFiles);
            _cancellationTokenSource = new CancellationTokenSource();
            SetRunningState(true);

            var diffBag = new ConcurrentBag<DiffResult>();
            var processed = 0;
            long matchedFiles = 0;
            long missingFiles = 0;
            long diffCount = 0;
            long compareErrorCount = 0;
            var token = _cancellationTokenSource.Token;
            var wasCancelled = false;
            LogInfo($"並列比較開始 MaxDegreeOfParallelism={Environment.ProcessorCount}");

            try
            {
                await Task.Run(() =>
                {
                    Parallel.ForEach(
                        sourceFiles,
                        new ParallelOptions
                        {
                            CancellationToken = token,
                            MaxDegreeOfParallelism = Environment.ProcessorCount
                        },
                        sourceFile =>
                        {
                            token.ThrowIfCancellationRequested();
                            DiffResult diff = null;
                            FileItem destinationFile = null;
                            try
                            {
                                if (destinationMap.TryGetValue(sourceFile.RelativePath, out destinationFile))
                                {
                                    Interlocked.Increment(ref matchedFiles);
                                    if (FilesAreDifferent(sourceFile, destinationFile))
                                    {
                                        diff = CreateDiffResult(sourceFile, destinationFile);
                                        Interlocked.Increment(ref diffCount);
                                    }
                                }
                                else
                                {
                                    Interlocked.Increment(ref missingFiles);
                                }
                            }
                            catch (Exception compareEx)
                            {
                                Interlocked.Increment(ref compareErrorCount);
                                LogError($"ファイル比較中にエラーが発生しました。Source='{sourceFile.FullPath}' Destination='{destinationFile?.FullPath ?? "N/A"}'", compareEx);
                            }
                            if (diff != null)
                            {
                                diffBag.Add(diff);
                            }

                            var current = Interlocked.Increment(ref processed);
                            ReportProgress(current, total);
                        });
                }, token);
            }
            catch (OperationCanceledException)
            {
                // 中止要求なので握りつぶす
                wasCancelled = true;
                LogInfo("比較処理がキャンセルされました。");
            }
            catch (Exception ex)
            {
                LogError("比較処理でエラーが発生しました。", ex);
                MessageBox.Show($"比較処理でエラーが発生しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetRunningState(false);
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }
            LogInfo($"比較処理終了 Processed={processed}/{total} DiffCount={diffCount} Matched={matchedFiles} Missing={missingFiles} CompareErrors={compareErrorCount} Cancelled={wasCancelled}");

            var orderedResults = diffBag
                .OrderBy(r => r.SourceDirectory, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.SourceFileName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var result in orderedResults)
            {
                _diffResults.Add(result);
            }
            LogInfo($"結果表示完了 件数={orderedResults.Count}");

            UpdateProgressLabel(processed, total);
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!_isRunning)
            {
                return;
            }
            LogInfo("中止ボタン押下: キャンセル要求を発行します。");
            _cancellationTokenSource?.Cancel();
        }

        private async void btnRunWinMerge_Click(object sender, EventArgs e)
        {
            if (_isRunning || _isWinMergeComparisonRunning)
            {
                MessageBox.Show("比較処理中はWinMergeによる再確認は実行できません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var winMergePath = txtWinMergePath.Text.Trim();
            if (string.IsNullOrWhiteSpace(winMergePath) || !File.Exists(winMergePath))
            {
                MessageBox.Show("WinMergeのパスが正しく設定されていません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_diffResults.Count == 0)
            {
                MessageBox.Show("DataGridViewに再確認する行がありません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _isWinMergeComparisonRunning = true;
            UpdateWinMergeButtonState();
            var originalCursor = Cursor;
            Cursor = Cursors.WaitCursor;
            LogInfo($"WinMerge再比較開始 件数={_diffResults.Count}");

            try
            {
                var snapshot = _diffResults.ToList();
                var removed = 0;
                var failures = 0;
                var skipped = 0;
                foreach (var diff in snapshot)
                {
                    if (!diff.CanOpenWinMerge)
                    {
                        skipped++;
                        continue;
                    }

                    int exitCode;
                    try
                    {
                        exitCode = await CompareWithWinMergeAsync(winMergePath, diff);
                    }
                    catch (Exception ex)
                    {
                        failures++;
                        LogError($"WinMerge比較中にエラーが発生しました。Source='{diff.SourceFullPath}' Destination='{diff.DestinationFullPath}'", ex);
                        continue;
                    }

                    if (exitCode == 0)
                    {
                        _diffResults.Remove(diff);
                        removed++;
                    }
                    else if (exitCode == 1)
                    {
                        // 差分ありのままなので何もしない
                    }
                    else
                    {
                        failures++;
                        LogInfo($"WinMergeが想定外の終了コード {exitCode} を返しました。Source='{diff.SourceFullPath}' Destination='{diff.DestinationFullPath}'");
                    }
                }

                LogInfo($"WinMerge再比較完了 Removed={removed} Failures={failures} Skipped={skipped} Remaining={_diffResults.Count}");
                var message = $"WinMergeでの再比較が完了しました。\n削除: {removed}件\n失敗: {failures}件\nスキップ: {skipped}件\n残件: {_diffResults.Count}件";
                MessageBox.Show(message, "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                _isWinMergeComparisonRunning = false;
                UpdateWinMergeButtonState();
                Cursor = originalCursor;
            }
        }

        private void dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _diffResults.Count)
            {
                return;
            }

            LaunchWinMerge(_diffResults[e.RowIndex]);
        }

        private void dataGridView_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                e.ToolTipText = string.Empty;
                return;
            }

            var value = dataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            e.ToolTipText = value?.ToString() ?? string.Empty;
        }

        private void dataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                dataGridView.SelectAll();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedRowsToClipboard();
                e.Handled = true;
            }
        }

        private void CopySelectedRowsToClipboard()
        {
            var rows = dataGridView.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Selected && !r.IsNewRow)
                .OrderBy(r => r.Index)
                .ToList();

            if (rows.Count == 0)
            {
                return;
            }

            var columnHeaders = dataGridView.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible)
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            var sb = new StringBuilder();
            sb.AppendLine(string.Join("\t", columnHeaders.Select(c => c.HeaderText)));

            foreach (var row in rows)
            {
                var values = columnHeaders
                    .Select(c => row.Cells[c.Index].Value?.ToString() ?? string.Empty);
                sb.AppendLine(string.Join("\t", values));
            }

            try
            {
                Clipboard.SetText(sb.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリップボードへのコピーに失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void LaunchWinMerge(DiffResult diff)
        {
            var winMergePath = txtWinMergePath.Text.Trim();
            if (string.IsNullOrWhiteSpace(winMergePath) || !File.Exists(winMergePath))
            {
                MessageBox.Show("WinMergeのパスが正しく設定されていません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!diff.CanOpenWinMerge)
            {
                MessageBox.Show("比較対象のファイルが存在しないためWinMergeを起動できません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = winMergePath,
                    Arguments = $"\"{diff.SourceFullPath}\" \"{diff.DestinationFullPath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WinMergeの起動に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Task<int> CompareWithWinMergeAsync(string winMergePath, DiffResult diff)
        {
            return Task.Run(() =>
            {
                var psi = new ProcessStartInfo
                {
                    FileName = winMergePath,
                    Arguments = BuildWinMergeArguments(diff),
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (var process = Process.Start(psi))
                {
                    if (process == null)
                    {
                        throw new InvalidOperationException("WinMergeの起動に失敗しました。");
                    }

                    process.WaitForExit();
                    return process.ExitCode;
                }
            });
        }

        private static string BuildWinMergeArguments(DiffResult diff)
        {
            return $"-minimize /noninteractive /quickcompare \"{diff.SourceFullPath}\" \"{diff.DestinationFullPath}\"";
        }


        private void SetRunningState(bool isRunning)
        {
            _isRunning = isRunning;
            btnStart.Enabled = !isRunning;
            btnStop.Enabled = isRunning;
            btnRefDirSrc.Enabled = !isRunning;
            btnRefDirDst.Enabled = !isRunning;
            btnRefWinMerge.Enabled = !isRunning;
            txtDirPathSrc.Enabled = !isRunning;
            txtDirPathDst.Enabled = !isRunning;
            txtWinMergePath.Enabled = !isRunning;
            chkEnableSubDir.Enabled = !isRunning;
            UpdateWinMergeButtonState();
            Cursor = isRunning ? Cursors.WaitCursor : Cursors.Default;
        }

        private void UpdateWinMergeButtonState()
        {
            if (btnRunWinMerge == null)
            {
                return;
            }

            btnRunWinMerge.Enabled = !_isRunning && !_isWinMergeComparisonRunning;
        }

        private void ReportProgress(int processed, int total)
        {
            if (IsDisposed)
            {
                return;
            }

            var percent = CalculatePercent(processed, total);
            var lastPercent = Volatile.Read(ref _lastUiProgressPercent);
            if (percent == lastPercent && processed != total)
            {
                return;
            }
            Volatile.Write(ref _lastUiProgressPercent, percent);

            if (InvokeRequired)
            {
                BeginInvoke((Action)(() => UpdateProgressLabel(processed, total)));
            }
            else
            {
                UpdateProgressLabel(processed, total);
            }
        }

        private void UpdateProgressLabel(int processed, int total)
        {
            if (total <= 0)
            {
                total = 1;
            }

            if (processed > total)
            {
                processed = total;
            }

            progressBar.Maximum = total;
            progressBar.Value = Math.Min(processed, total);
            lblInfo.Text = $"{processed}/{total}ファイル";
            LogProgressIfNeeded(processed, total);
        }

        private void SelectFolder(TextBox target)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "フォルダを選択してください";
                dialog.ShowNewFolderButton = false;
                if (Directory.Exists(target.Text))
                {
                    dialog.SelectedPath = target.Text;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    target.Text = dialog.SelectedPath;
                }
            }
        }

        private static List<FileItem> EnumerateFileItems(string rootPath, SearchOption searchOption)
        {
            var absoluteRoot = Path.GetFullPath(rootPath);
            var normalizedRoot = EnsureTrailingSeparator(absoluteRoot);

            return Directory.EnumerateFiles(absoluteRoot, "*", searchOption)
                .Select(path =>
                {
                    var info = new FileInfo(path);
                    return new FileItem
                    {
                        FullPath = path,
                        DirectoryPath = info.DirectoryName ?? absoluteRoot,
                        FileName = info.Name,
                        RelativePath = GetRelativePath(normalizedRoot, path),
                        Length = info.Length
                    };
                })
                .ToList();
        }

        private static Dictionary<string, FileItem> BuildFileMap(IEnumerable<FileItem> files)
        {
            var map = new Dictionary<string, FileItem>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in files)
            {
                map[file.RelativePath] = file;
            }
            return map;
        }

        private static bool FilesAreDifferent(FileItem source, FileItem destination)
        {
            if (source.Length != destination.Length)
            {
                return true;
            }

            const int bufferSize = 64 * 1024;
            var bufferSource = new byte[bufferSize];
            var bufferDestination = new byte[bufferSize];

            using (var sourceStream = new FileStream(source.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var destinationStream = new FileStream(destination.FullPath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                while (true)
                {
                    var readSource = sourceStream.Read(bufferSource, 0, bufferSize);
                    var readDestination = destinationStream.Read(bufferDestination, 0, bufferSize);
                    if (readSource != readDestination)
                    {
                        return true;
                    }

                    if (readSource == 0)
                    {
                        break;
                    }

                    for (var i = 0; i < readSource; i++)
                    {
                        if (bufferSource[i] != bufferDestination[i])
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private static DiffResult CreateDiffResult(FileItem source, FileItem destination)
        {
            return new DiffResult
            {
                SourceDirectory = source.DirectoryPath,
                SourceFileName = source.FileName,
                DestinationDirectory = destination.DirectoryPath,
                DestinationFileName = destination.FileName,
                SourceFullPath = source.FullPath,
                DestinationFullPath = destination.FullPath
            };
        }

        private static string GetRelativePath(string normalizedRoot, string fullPath)
        {
            var root = EnsureTrailingSeparator(Path.GetFullPath(normalizedRoot));
            var normalizedFullPath = Path.GetFullPath(fullPath);
            if (normalizedFullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                return normalizedFullPath.Substring(root.Length);
            }

            return Path.GetFileName(normalizedFullPath);
        }

        private static string EnsureTrailingSeparator(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return path;
            }

            if (!path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                path += Path.DirectorySeparatorChar;
            }

            return path;
        }

        private void EnablePathDragDrop(TextBox textBox, Func<string, bool> validator)
        {
            if (textBox == null || validator == null)
            {
                return;
            }

            textBox.AllowDrop = true;
            textBox.DragEnter += (sender, e) =>
            {
                e.Effect = TryGetDroppedPath(e.Data, validator, out _)
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            };
            textBox.DragDrop += (sender, e) =>
            {
                if (TryGetDroppedPath(e.Data, validator, out var path))
                {
                    textBox.Text = path;
                }
            };
        }

        private static bool TryGetDroppedPath(IDataObject data, Func<string, bool> validator, out string path)
        {
            path = string.Empty;
            if (data == null || validator == null)
            {
                return false;
            }

            if (!data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            if (data.GetData(DataFormats.FileDrop) is string[] paths)
            {
                foreach (var candidate in paths)
                {
                    var trimmed = candidate?.Trim();
                    if (!string.IsNullOrEmpty(trimmed) && validator(trimmed))
                    {
                        path = trimmed;
                        return true;
                    }
                }
            }

            return false;
        }

        private bool ValidateDirectory(string path, string caption)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                LogInfo($"{caption}が未指定です。");
                MessageBox.Show($"{caption}を指定してください。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (!Directory.Exists(path))
            {
                LogInfo($"{caption} '{path}' が存在しません。");
                MessageBox.Show($"{caption}が存在しません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return true;
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsPath))
            {
                return;
            }

            try
            {
                using (var stream = File.OpenRead(_settingsPath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    if (serializer.ReadObject(stream) is AppSettings settings)
                    {
                        txtWinMergePath.Text = settings.WinMergePath ?? string.Empty;
                        txtDirPathSrc.Text = settings.SourceDirectory ?? string.Empty;
                        txtDirPathDst.Text = settings.DestinationDirectory ?? string.Empty;
                        chkEnableSubDir.Checked = settings.IncludeSubDirectories;
                        chkUseDebugLog.Checked = settings.EnableDebugLog;

                        if (settings.Width > 0 && settings.Height > 0)
                        {
                            Size = new Size(settings.Width, settings.Height);
                        }

                        if (settings.Left >= 0 && settings.Top >= 0)
                        {
                            StartPosition = FormStartPosition.Manual;
                            Location = new Point(settings.Left, settings.Top);
                        }

                        if (!string.IsNullOrWhiteSpace(settings.WindowState) &&
                            Enum.TryParse(settings.WindowState, out FormWindowState state))
                        {
                            WindowState = state;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("設定ファイルの読み込みに失敗しました。", ex);
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SaveSettings()
        {
            try
            {
                var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                var settings = new AppSettings
                {
                    WinMergePath = txtWinMergePath.Text,
                    SourceDirectory = txtDirPathSrc.Text,
                    DestinationDirectory = txtDirPathDst.Text,
                    IncludeSubDirectories = chkEnableSubDir.Checked,
                    EnableDebugLog = chkUseDebugLog.Checked,
                    Width = bounds.Width,
                    Height = bounds.Height,
                    Left = bounds.Left,
                    Top = bounds.Top,
                    WindowState = WindowState.ToString()
                };

                var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                using (var stream = File.Create(_settingsPath))
                {
                    serializer.WriteObject(stream, settings);
                }
            }
            catch (Exception ex)
            {
                LogError("設定ファイルの保存に失敗しました。", ex);
                MessageBox.Show($"設定ファイルの保存に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        [DataContract]
        private sealed class AppSettings
        {
            [DataMember] public string WinMergePath { get; set; } = string.Empty;
            [DataMember] public string SourceDirectory { get; set; } = string.Empty;
            [DataMember] public string DestinationDirectory { get; set; } = string.Empty;
            [DataMember] public bool IncludeSubDirectories { get; set; } = true;
            [DataMember] public bool EnableDebugLog { get; set; } = true;
            [DataMember] public int Width { get; set; }
            [DataMember] public int Height { get; set; }
            [DataMember] public int Left { get; set; } = -1;
            [DataMember] public int Top { get; set; } = -1;
            [DataMember] public string WindowState { get; set; } = FormWindowState.Normal.ToString();
        }

        private sealed class DiffResult
        {
            public string SourceDirectory { get; set; } = string.Empty;
            public string SourceFileName { get; set; } = string.Empty;
            public string DestinationDirectory { get; set; } = string.Empty;
            public string DestinationFileName { get; set; } = string.Empty;
            public string SourceFullPath { get; set; } = string.Empty;
            public string DestinationFullPath { get; set; } = string.Empty;

            public bool CanOpenWinMerge =>
                File.Exists(SourceFullPath) &&
                File.Exists(DestinationFullPath);
        }

        private sealed class FileItem
        {
            public string FullPath { get; set; } = string.Empty;
            public string DirectoryPath { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public string RelativePath { get; set; } = string.Empty;
            public long Length { get; set; }
        }

        private void chkUseDebugLog_CheckedChanged(object sender, EventArgs e)
        {
            UpdateLogEnabledState();
        }

        private void UpdateLogEnabledState()
        {
            _isLogEnabled = chkUseDebugLog.Checked;
        }

        private void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        private void LogError(string message, Exception exception)
        {
            var builder = new StringBuilder();
            builder.Append(message);
            if (exception != null)
            {
                builder.Append(" Exception=");
                builder.Append(exception);
            }
            WriteLog("ERROR", builder.ToString());
        }

        private void WriteLog(string level, string message)
        {
            if (!_isLogEnabled)
            {
                return;
            }

            try
            {
                var line = $"{DateTime.Now:yyyy/MM/dd HH:mm:ss.fff} [{level}] [T{Thread.CurrentThread.ManagedThreadId}] {message}";
                lock (_logLock)
                {
                    File.AppendAllText(_logPath, line + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // ログ出力に失敗してもアプリの処理は継続させる
            }
        }

        private void LogProgressIfNeeded(int processed, int total)
        {
            if (total <= 0)
            {
                return;
            }

            var percent = CalculatePercent(processed, total);
            if (percent < 0 || percent > 100)
            {
                return;
            }

            if (percent == _lastLoggedProgressPercent)
            {
                return;
            }

            if (percent % 10 == 0 || processed == total)
            {
                _lastLoggedProgressPercent = percent;
                LogInfo($"進捗 {processed}/{total} ({percent}%)");
            }
        }

        private static int CalculatePercent(int processed, int total)
        {
            if (total <= 0)
            {
                return 0;
            }

            var percent = (int)Math.Floor((double)processed / total * 100);
            if (percent < 0)
            {
                return 0;
            }

            if (percent > 100)
            {
                return 100;
            }

            return percent;
        }
    }
}
