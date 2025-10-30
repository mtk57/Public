using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    public partial class MethodListForm : Form
    {
        private const int MaxHistoryCount = 10;
        private readonly AppSettings _settings;

        public MethodListForm(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            InitializeComponent();
            InitializeControls();
            HookEvents();
            LoadSettings();
        }

        private void InitializeControls()
        {
            if (cmbExt.Items.Count > 0 && cmbExt.SelectedIndex < 0)
            {
                cmbExt.SelectedIndex = 0;
            }
        }

        private void HookEvents()
        {
            btnBrowse.Click += BtnBrowse_Click;
            btnCreate.Click += BtnCreate_Click;
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

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            CreateMethodList();
        }

        private void CreateMethodList()
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

            var loggingEnabled = IsLoggingEnabled;
            if (loggingEnabled)
            {
                MethodListLogger.LogInfo("==== メソッドリスト作成開始 ====");
                MethodListLogger.LogInfo($"対象フォルダ: {directoryPath}");
                MethodListLogger.LogInfo($"対象拡張子: {extension}");
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                var files = EnumerateSourceFiles(directoryPath, extension).ToList();
                if (loggingEnabled)
                {
                    MethodListLogger.LogInfo($"検出ファイル数: {files.Count}");
                }

                var results = new List<MethodDefinitionDetail>();
                foreach (var file in files)
                {
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
                    }
                    catch (JavaParseException ex)
                    {
                        if (loggingEnabled)
                        {
                            MethodListLogger.LogError($"解析失敗: {file} (行: {ex.LineNumber})");
                            if (!string.IsNullOrEmpty(ex.InvalidContent))
                            {
                                MethodListLogger.LogError($"問題箇所: {ex.InvalidContent}");
                            }

                            MethodListLogger.LogException(ex);
                            MethodListLogger.LogInfo("エラーのため処理を中断します。");
                        }

                        HandleJavaParseException(file, ex);
                        return;
                    }
                }

                var exportPath = BuildExportPath();
                WriteResults(exportPath, results);
                UpdateDirectoryHistory(directoryPath);
                SaveSettings();

                if (loggingEnabled)
                {
                    MethodListLogger.LogInfo($"出力ファイル: {exportPath}");
                    MethodListLogger.LogInfo($"出力件数: {results.Count}");
                    MethodListLogger.LogInfo("==== メソッドリスト作成完了 ====");
                }

                MessageBox.Show(this, $"メソッドリストを出力しました。\n{exportPath}", "結果",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenExportFolder(exportPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"メソッドリストの作成に失敗しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorLogger.LogException(ex);
                if (loggingEnabled)
                {
                    MethodListLogger.LogError($"メソッドリスト作成中にエラーが発生しました: {ex.Message}");
                    MethodListLogger.LogException(ex);
                }
            }
            finally
            {
                Cursor = Cursors.Default;
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
            var fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".tsv";
            return Path.Combine(baseDirectory, fileName);
        }

        private void WriteResults(string exportPath, IEnumerable<MethodDefinitionDetail> results)
        {
            if (string.IsNullOrEmpty(exportPath))
            {
                throw new ArgumentException("出力先パスが不正です。", nameof(exportPath));
            }

            var headers = new[]
            {
                "FilePath",
                "FileName",
                "PackageName",
                "ClassName",
                "MethodSignature"
            };

            using (var writer = new StreamWriter(exportPath, false, Encoding.UTF8))
            {
                writer.WriteLine(string.Join("\t", headers));
                foreach (var detail in results ?? Enumerable.Empty<MethodDefinitionDetail>())
                {
                    var fields = new[]
                    {
                        EscapeForTsv(detail?.FilePath),
                        EscapeForTsv(detail?.FileName),
                        EscapeForTsv(detail?.PackageName),
                        EscapeForTsv(detail?.ClassName),
                        EscapeForTsv(detail?.MethodSignature)
                    };
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
                }

                SettingsManager.Save(_settings);
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
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
            if (IsLoggingEnabled)
            {
                MethodListLogger.LogError(message.TrimEnd());
                MethodListLogger.LogException(ex);
            }
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
            SaveSettings();
        }

        private bool IsLoggingEnabled
        {
            get { return chkEnableLogging != null && chkEnableLogging.Checked; }
        }
    }
}
