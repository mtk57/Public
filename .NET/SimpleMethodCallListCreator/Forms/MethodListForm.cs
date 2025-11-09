using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    public enum MethodListExportMode
    {
        Standard,
        RowNumber
    }

    public partial class MethodListForm : Form
    {
        private const int MaxHistoryCount = 10;
        private readonly AppSettings _settings;
        private readonly MethodListExportMode _exportMode;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isProcessing;

        public MethodListForm(AppSettings settings, MethodListExportMode exportMode = MethodListExportMode.Standard)
        {
            _settings = settings ?? new AppSettings();
            _exportMode = exportMode;
            InitializeComponent();
            InitializeControls();
            HookEvents();
            LoadSettings();
            UpdateTitleForMode();
        }

        private void InitializeControls()
        {
            if (cmbExt.Items.Count > 0 && cmbExt.SelectedIndex < 0)
            {
                cmbExt.SelectedIndex = 0;
            }

            pbProgress.Minimum = 0;
            pbProgress.Maximum = 1;
            pbProgress.Value = 0;
            pbProgress.Visible = false;
            btnCancel.Enabled = false;
            lblProgressStatus.Text = string.Empty;
        }

        private void HookEvents()
        {
            btnBrowse.Click += BtnBrowse_Click;
            btnCreate.Click += BtnCreate_Click;
            btnCancel.Click += BtnCancel_Click;
            cmbDirPath.DragEnter += CmbDirPath_DragEnter;
            cmbDirPath.DragDrop += CmbDirPath_DragDrop;
            FormClosing += MethodListForm_FormClosing;
        }

        private void LoadSettings()
        {
            if (_settings.RecentMethodListDirectories == null)
            {
                _settings.RecentMethodListDirectories = new List<string>();
            }

            cmbDirPath.Items.Clear();
            cmbDirPath.Items.AddRange(_settings.RecentMethodListDirectories.ToArray());

            if (_settings.SelectedMethodListDirectoryIndex >= 0 &&
                _settings.SelectedMethodListDirectoryIndex < cmbDirPath.Items.Count)
            {
                cmbDirPath.SelectedIndex = _settings.SelectedMethodListDirectoryIndex;
            }
            else if (!string.IsNullOrEmpty(_settings.LastMethodListDirectory))
            {
                var index = cmbDirPath.FindStringExact(_settings.LastMethodListDirectory);
                if (index >= 0)
                {
                    cmbDirPath.SelectedIndex = index;
                }
                else
                {
                    cmbDirPath.SelectedIndex = -1;
                    cmbDirPath.Text = _settings.LastMethodListDirectory;
                }
            }
            else
            {
                cmbDirPath.SelectedIndex = -1;
                cmbDirPath.Text = string.Empty;
            }

            if (cmbExt.Items.Count > 0)
            {
                var selectedIndex = _settings.SelectedMethodListExtensionIndex;
                if (selectedIndex >= 0 && selectedIndex < cmbExt.Items.Count)
                {
                    cmbExt.SelectedIndex = selectedIndex;
                }
            }

            chkEnableLogging.Checked = _settings.EnableMethodListLogging;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "メソッドリストを作成するフォルダを選択してください。";
                dialog.ShowNewFolderButton = false;
                var currentPath = (cmbDirPath.Text ?? string.Empty).Trim();
                if (Directory.Exists(currentPath))
                {
                    dialog.SelectedPath = currentPath;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    cmbDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private async void BtnCreate_Click(object sender, EventArgs e)
        {
            if (_isProcessing)
            {
                return;
            }

            await StartMethodListCreationAsync();
        }

        private async Task StartMethodListCreationAsync()
        {
            var directoryPath = (cmbDirPath.Text ?? string.Empty).Trim();
            if (directoryPath.Length == 0)
            {
                MessageBox.Show(this, "対象フォルダパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(directoryPath))
            {
                MessageBox.Show(this, "指定されたフォルダが見つかりません。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbExt.SelectedItem == null)
            {
                MessageBox.Show(this, "対象拡張子を選択してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var extension = ResolveExtension(cmbExt.SelectedItem.ToString());
            if (extension == null)
            {
                MessageBox.Show(this, "未対応の拡張子です。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var files = EnumerateSourceFiles(directoryPath, extension).ToList();
            var loggingEnabled = IsLoggingEnabled;

            if (files.Count == 0)
            {
                MessageBox.Show(this, "対象フォルダに解析対象のファイルがありません。", "情報",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (loggingEnabled)
                {
                    MethodListLogger.LogInfo("解析対象ファイルが見つかりませんでした。処理を終了します。");
                }

                return;
            }

            if (loggingEnabled)
            {
                MethodListLogger.LogInfo("==== メソッドリスト作成開始 ====");
                MethodListLogger.LogInfo($"対象フォルダ: {directoryPath}");
                MethodListLogger.LogInfo($"対象拡張子: {extension}");
                MethodListLogger.LogInfo($"検出ファイル数: {files.Count}");
            }

            PrepareProcessingState(files.Count);

            _cancellationTokenSource = new CancellationTokenSource();
            var token = _cancellationTokenSource.Token;
            var progress = new Progress<MethodListProgressStatus>(status =>
            {
                UpdateProgressStatus(status);
            });

            MethodListProcessResult result = null;
            Cursor = Cursors.WaitCursor;
            try
            {
                result = await Task.Run(
                    () => ProcessMethodList(files, loggingEnabled, token, progress),
                    token);

                UpdateDirectoryHistory(directoryPath);
                SaveSettings();

                var completionMessage = new StringBuilder();
                completionMessage.AppendLine("メソッドリストを出力しました。");
                completionMessage.AppendLine(result.ExportPath);
                completionMessage.AppendLine($"成功: {result.SuccessCount}  失敗: {result.FailureCount}");

                MessageBox.Show(this, completionMessage.ToString(), "結果",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenExportFolder(result.ExportPath);

                if (result.FailureCount > 0)
                {
                    var failureMessage = new StringBuilder();
                    failureMessage.AppendLine("一部のファイルで解析に失敗しました。");
                    failureMessage.AppendLine($"失敗件数: {result.FailureCount}");
                    failureMessage.AppendLine("失敗ファイル:");

                    var displayCount = Math.Min(10, result.FailedFiles.Count);
                    for (var i = 0; i < displayCount; i++)
                    {
                        failureMessage.AppendLine($"  - {result.FailedFiles[i]}");
                    }

                    if (result.FailedFiles.Count > displayCount)
                    {
                        failureMessage.AppendLine("  ...（残りは methodlist.log を参照してください）");
                    }

                    MessageBox.Show(this, failureMessage.ToString(), "解析失敗", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (OperationCanceledException)
            {
                if (loggingEnabled)
                {
                    MethodListLogger.LogInfo("処理はユーザーにより中断されました。");
                    MethodListLogger.LogInfo("==== メソッドリスト作成中断 ====");
                }

                MessageBox.Show(this, "メソッドリストの作成を中断しました。", "情報",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"メソッドリストの作成に失敗しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorLogger.LogException(ex);
                MethodListLogger.LogError($"メソッドリスト作成中にエラーが発生しました: {ex.Message}");
                MethodListLogger.LogException(ex);
                if (loggingEnabled)
                {
                    MethodListLogger.LogInfo("==== メソッドリスト作成失敗 ====");
                }
            }
            finally
            {
                Cursor = Cursors.Default;
                ResetProcessingState();
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;

                if (result != null && loggingEnabled)
                {
                    MethodListLogger.LogInfo($"出力ファイル: {result.ExportPath}");
                    MethodListLogger.LogInfo($"成功件数: {result.SuccessCount}, 失敗件数: {result.FailureCount}");
                    if (result.FailureCount > 0)
                    {
                        MethodListLogger.LogInfo("一部のファイルで解析に失敗しました。詳細は上記ログを参照してください。");
                    }
                    MethodListLogger.LogInfo("==== メソッドリスト作成完了 ====");
                }
            }
        }

        private IEnumerable<string> EnumerateSourceFiles(string directoryPath, string extension)
        {
            return Directory.EnumerateFiles(directoryPath, "*" + extension, SearchOption.AllDirectories);
        }

        private string ResolveExtension(string selected)
        {
            if (string.Equals(selected, "Java", StringComparison.OrdinalIgnoreCase))
            {
                return ".java";
            }

            return null;
        }

        private string BuildExportPath()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
            var suffix = _exportMode == MethodListExportMode.RowNumber ? "_row" : string.Empty;
            var fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + suffix + ".tsv";
            return Path.Combine(baseDirectory, fileName);
        }

        private void WriteResults(string exportPath, IEnumerable<MethodDefinitionDetail> results)
        {
            if (string.IsNullOrEmpty(exportPath))
            {
                throw new ArgumentException("出力先パスが不正です。", nameof(exportPath));
            }

            var headers = new List<string>
            {
                "FilePath",
                "FileName",
                "PackageName",
                "ClassName",
                "MethodSignature"
            };
            if (_exportMode == MethodListExportMode.RowNumber)
            {
                headers.Add("RowNum");
            }

            using (var writer = new StreamWriter(exportPath, false, Encoding.UTF8))
            {
                writer.WriteLine(string.Join("\t", headers));
                foreach (var detail in results ?? Enumerable.Empty<MethodDefinitionDetail>())
                {
                    var fields = new List<string>
                    {
                        EscapeForTsv(detail?.FilePath),
                        EscapeForTsv(detail?.FileName),
                        EscapeForTsv(detail?.PackageName),
                        EscapeForTsv(detail?.ClassName),
                        EscapeForTsv(detail?.MethodSignature)
                    };
                    if (_exportMode == MethodListExportMode.RowNumber)
                    {
                        var lineNumber = detail?.LineNumber > 0
                            ? detail.LineNumber.ToString(CultureInfo.InvariantCulture)
                            : string.Empty;
                        fields.Add(lineNumber);
                    }
                    writer.WriteLine(string.Join("\t", fields));
                }
            }
        }

        private string EscapeForTsv(string value)
        {
            return (value ?? string.Empty).Replace("\t", "    ");
        }

        private void UpdateDirectoryHistory(string directoryPath)
        {
            if (_settings.RecentMethodListDirectories == null)
            {
                _settings.RecentMethodListDirectories = new List<string>();
            }

            UpdateHistoryList(_settings.RecentMethodListDirectories, directoryPath, true);

            cmbDirPath.BeginUpdate();
            try
            {
                cmbDirPath.Items.Clear();
                cmbDirPath.Items.AddRange(_settings.RecentMethodListDirectories.ToArray());
                cmbDirPath.Text = directoryPath;
            }
            finally
            {
                cmbDirPath.EndUpdate();
            }

            _settings.LastMethodListDirectory = directoryPath;
            _settings.SelectedMethodListDirectoryIndex = cmbDirPath.FindStringExact(directoryPath);
            _settings.SelectedMethodListExtensionIndex = cmbExt.SelectedIndex;
        }

        private void UpdateHistoryList(List<string> history, string value, bool caseInsensitive)
        {
            if (history == null || string.IsNullOrEmpty(value))
            {
                return;
            }

            var comparer = caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            for (var i = history.Count - 1; i >= 0; i--)
            {
                if (comparer.Equals(history[i], value))
                {
                    history.RemoveAt(i);
                }
            }

            history.Insert(0, value);
            if (history.Count > MaxHistoryCount)
            {
                history.RemoveRange(MaxHistoryCount, history.Count - MaxHistoryCount);
            }
        }

        private void SaveSettings()
        {
            try
            {
                if (_settings != null)
                {
                    _settings.EnableMethodListLogging = chkEnableLogging.Checked;

                    var directoryHistory = new List<string>();
                    foreach (var item in cmbDirPath.Items)
                    {
                        var text = (item as string ?? string.Empty).Trim();
                        if (text.Length == 0)
                        {
                            continue;
                        }

                        if (!directoryHistory.Exists(
                                d => string.Equals(d, text, StringComparison.OrdinalIgnoreCase)))
                        {
                            directoryHistory.Add(text);
                        }
                    }

                    var currentDirectory = (cmbDirPath.Text ?? string.Empty).Trim();
                    if (currentDirectory.Length > 0 &&
                        !directoryHistory.Exists(
                            d => string.Equals(d, currentDirectory, StringComparison.OrdinalIgnoreCase)))
                    {
                        directoryHistory.Insert(0, currentDirectory);
                    }

                    if (directoryHistory.Count > MaxHistoryCount)
                    {
                        directoryHistory.RemoveRange(MaxHistoryCount, directoryHistory.Count - MaxHistoryCount);
                    }

                    _settings.RecentMethodListDirectories = directoryHistory;
                    _settings.LastMethodListDirectory = currentDirectory;
                    _settings.SelectedMethodListDirectoryIndex = directoryHistory.FindIndex(
                        d => string.Equals(d, currentDirectory, StringComparison.OrdinalIgnoreCase));
                    _settings.SelectedMethodListExtensionIndex = cmbExt.SelectedIndex;
                }

                SettingsManager.Save(_settings);
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (!_isProcessing || _cancellationTokenSource == null)
            {
                return;
            }

            btnCancel.Enabled = false;
            _cancellationTokenSource.Cancel();
        }

        private void PrepareProcessingState(int fileCount)
        {
            _isProcessing = true;
            btnCreate.Enabled = false;
            btnCancel.Enabled = true;

            pbProgress.Minimum = 0;
            pbProgress.Maximum = Math.Max(1, fileCount);
            pbProgress.Value = 0;
            pbProgress.Style = ProgressBarStyle.Continuous;
            pbProgress.Visible = true;
            pbProgress.Refresh();

            UpdateProgressStatus(new MethodListProgressStatus(fileCount, 0, 0, 0));
        }

        private void ResetProcessingState()
        {
            _isProcessing = false;
            btnCreate.Enabled = true;
            btnCancel.Enabled = false;
            pbProgress.Visible = false;
            pbProgress.Value = 0;
            pbProgress.Maximum = 1;
            lblProgressStatus.Text = string.Empty;
        }

        private MethodListProcessResult ProcessMethodList(
            List<string> files,
            bool loggingEnabled,
            CancellationToken token,
            IProgress<MethodListProgressStatus> progress)
        {
            var total = files.Count;
            var processed = 0;
            var successCount = 0;
            var failureCount = 0;
            var results = new List<MethodDefinitionDetail>();
            var failedFiles = new List<string>();

            foreach (var file in files)
            {
                token.ThrowIfCancellationRequested();

                try
                {
                    if (loggingEnabled)
                    {
                        MethodListLogger.LogInfo($"解析開始: {file}");
                    }

                    var definitions = JavaMethodCallAnalyzer.ExtractMethodDefinitions(file);
                    results.AddRange(definitions);

                    if (loggingEnabled)
                    {
                        MethodListLogger.LogInfo($"解析成功: {file} (メソッド定義数: {definitions.Count})");
                    }

                    successCount++;
                }
                catch (JavaParseException ex)
                {
                    MethodListLogger.LogError($"解析失敗: {file} (行: {ex.LineNumber})");
                    if (!string.IsNullOrEmpty(ex.InvalidContent))
                    {
                        MethodListLogger.LogError($"問題箇所: {ex.InvalidContent}");
                    }

                    MethodListLogger.LogException(ex);
                    failedFiles.Add($"{file} (行: {ex.LineNumber})");
                    failureCount++;
                }
                catch (Exception ex)
                {
                    MethodListLogger.LogError($"解析失敗(例外): {file} メッセージ: {ex.Message}");
                    MethodListLogger.LogException(ex);
                    failedFiles.Add($"{file} ({ex.Message})");
                    failureCount++;
                }

                processed++;
                progress?.Report(new MethodListProgressStatus(total, processed, successCount, failureCount));
            }

            token.ThrowIfCancellationRequested();

            var exportPath = BuildExportPath();
            WriteResults(exportPath, results);

            return new MethodListProcessResult(exportPath, successCount, failureCount, failedFiles);
        }

        private sealed class MethodListProcessResult
        {
            public MethodListProcessResult(string exportPath, int successCount, int failureCount, List<string> failedFiles)
            {
                ExportPath = exportPath;
                SuccessCount = successCount;
                FailureCount = failureCount;
                FailedFiles = failedFiles ?? new List<string>();
            }

            public string ExportPath { get; }
            public int SuccessCount { get; }
            public int FailureCount { get; }
            public List<string> FailedFiles { get; }
        }

        private sealed class MethodListProgressStatus
        {
            public MethodListProgressStatus(int total, int processed, int success, int failure)
            {
                Total = Math.Max(0, total);
                Processed = Math.Max(0, processed);
                Success = Math.Max(0, success);
                Failure = Math.Max(0, failure);
            }

            public int Total { get; }
            public int Processed { get; }
            public int Success { get; }
            public int Failure { get; }
        }

        private void UpdateProgressStatus(MethodListProgressStatus status)
        {
            if (status == null)
            {
                return;
            }

            var total = Math.Max(1, status.Total);
            pbProgress.Maximum = total;
            var value = Math.Max(pbProgress.Minimum, Math.Min(total, status.Processed));
            pbProgress.Value = value;

            lblProgressStatus.Text =
                $"総数: {status.Total} / 処理済: {status.Processed} (成功: {status.Success}, 失敗: {status.Failure})";
            lblProgressStatus.Refresh();
        }

        private void HandleJavaParseException(string filePath, JavaParseException ex)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Javaファイルの解析に失敗しました。");
            builder.AppendLine($"ファイル: {filePath}");
            builder.AppendLine($"行番号: {ex.LineNumber}");
            if (!string.IsNullOrEmpty(ex.InvalidContent))
            {
                builder.AppendLine($"内容: {ex.InvalidContent}");
            }

            var message = builder.ToString();
            MessageBox.Show(this, message, "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            ErrorLogger.LogError(message.TrimEnd());
            MethodListLogger.LogError(message.TrimEnd());
            MethodListLogger.LogException(ex);
        }

        private void CmbDirPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var items = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (items != null && items.Length > 0 && Directory.Exists(items[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }

            e.Effect = DragDropEffects.None;
        }

        private void CmbDirPath_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var items = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (items == null || items.Length == 0)
            {
                return;
            }

            var candidate = items[0];
            if (!Directory.Exists(candidate))
            {
                MessageBox.Show(this, "フォルダをドロップしてください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            cmbDirPath.Text = candidate;
        }

        private void OpenExportFolder(string exportPath)
        {
            try
            {
                var directory = Path.GetDirectoryName(exportPath);
                if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                {
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = directory,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
        }

        private void MethodListForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_isProcessing)
            {
                MessageBox.Show(this, "メソッドリスト作成中はウィンドウを閉じることができません。", "情報",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true;
                return;
            }

            SaveSettings();
        }

        private bool IsLoggingEnabled
        {
            get { return chkEnableLogging != null && chkEnableLogging.Checked; }
        }

        private void UpdateTitleForMode()
        {
            if (_exportMode == MethodListExportMode.RowNumber)
            {
                Text = "Method List (RowNumber)";
            }
        }
    }
}
