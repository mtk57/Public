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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleDiff
{
    public partial class MainForm : Form
    {
        private readonly string _settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleDiff.json");
        private readonly BindingList<DiffResult> _diffResults = new BindingList<DiffResult>();
        private CancellationTokenSource _cancellationTokenSource;
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
            dataGridView.CellDoubleClick += dataGridView_CellDoubleClick;
        }

        private void InitializeCustomComponents()
        {
            dataGridView.AutoGenerateColumns = false;
            dataGridView.AllowUserToAddRows = false;
            dataGridView.AllowUserToDeleteRows = false;
            dataGridView.MultiSelect = false;
            dataGridView.ReadOnly = true;
            dataGridView.RowHeadersVisible = false;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            clmDirPathSrc.DataPropertyName = nameof(DiffResult.SourceDirectory);
            clmFileNameSrc.DataPropertyName = nameof(DiffResult.SourceFileName);
            clmDirPathDst.DataPropertyName = nameof(DiffResult.DestinationDirectory);
            clmFileNameDst.DataPropertyName = nameof(DiffResult.DestinationFileName);
            dataGridView.DataSource = _diffResults;
            chkEnableSubDir.Checked = true;
            progressBar.Minimum = 0;
            progressBar.Value = 0;
            lblInfo.Text = "0/0ファイル";
            btnStop.Enabled = false;
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

            if (!ValidateDirectory(sourceRoot, "比較元フォルダパス"))
            {
                return;
            }

            if (!ValidateDirectory(destinationRoot, "比較先フォルダパス"))
            {
                return;
            }

            try
            {
                sourceRoot = Path.GetFullPath(sourceRoot);
                destinationRoot = Path.GetFullPath(destinationRoot);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"フォルダパスの解決に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var searchOption = chkEnableSubDir.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            List<FileItem> sourceFiles;
            List<FileItem> destinationFiles;

            try
            {
                sourceFiles = EnumerateFileItems(sourceRoot, searchOption);
                destinationFiles = EnumerateFileItems(destinationRoot, searchOption);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"フォルダの走査に失敗しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (sourceFiles.Count == 0)
            {
                MessageBox.Show("比較元フォルダにファイルが存在しません。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _diffResults.Clear();
            var total = sourceFiles.Count;
            progressBar.Maximum = total;
            progressBar.Value = 0;
            UpdateProgressLabel(0, total);

            var destinationMap = BuildFileMap(destinationFiles);
            _cancellationTokenSource = new CancellationTokenSource();
            SetRunningState(true);

            var diffBag = new ConcurrentBag<DiffResult>();
            var processed = 0;
            var token = _cancellationTokenSource.Token;

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

                            if (destinationMap.TryGetValue(sourceFile.RelativePath, out var destinationFile))
                            {
                                if (FilesAreDifferent(sourceFile, destinationFile))
                                {
                                    diff = CreateDiffResult(sourceFile, destinationFile);
                                }
                            }
                            else
                            {
                                diff = CreateMissingDestinationResult(sourceFile, destinationRoot);
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
            }
            catch (Exception ex)
            {
                MessageBox.Show($"比較処理でエラーが発生しました。\n{ex.Message}", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetRunningState(false);
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
            }

            var orderedResults = diffBag
                .OrderBy(r => r.SourceDirectory, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.SourceFileName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var result in orderedResults)
            {
                _diffResults.Add(result);
            }

            UpdateProgressLabel(processed, total);
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (!_isRunning)
            {
                return;
            }

            _cancellationTokenSource?.Cancel();
        }

        private void dataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _diffResults.Count)
            {
                return;
            }

            LaunchWinMerge(_diffResults[e.RowIndex]);
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
            Cursor = isRunning ? Cursors.WaitCursor : Cursors.Default;
        }

        private void ReportProgress(int processed, int total)
        {
            if (IsDisposed)
            {
                return;
            }

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

        private static DiffResult CreateMissingDestinationResult(FileItem source, string destinationRoot)
        {
            var expectedDestinationPath = Path.Combine(destinationRoot, source.RelativePath);
            var destinationDirectory = Path.GetDirectoryName(expectedDestinationPath) ?? destinationRoot;

            return new DiffResult
            {
                SourceDirectory = source.DirectoryPath,
                SourceFileName = source.FileName,
                DestinationDirectory = destinationDirectory,
                DestinationFileName = Path.GetFileName(expectedDestinationPath),
                SourceFullPath = source.FullPath,
                DestinationFullPath = expectedDestinationPath
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

        private bool ValidateDirectory(string path, string caption)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show($"{caption}を指定してください。", "SimpleDiff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (!Directory.Exists(path))
            {
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
    }
}
