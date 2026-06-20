using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading; // Interlockedのために追加
using System.Threading.Tasks;
using System.Windows.Forms;
using SimpleGrep.Core;

namespace SimpleGrep
{
    public partial class MainForm : Form
    {
        private const string SettingsFileName = "SimpleGrep.settings.json";
        private const int MaxHistoryCount = 10;
        private CancellationTokenSource searchCancellationTokenSource;
        private List<SearchResult> currentSearchResults = new List<SearchResult>();
        private string multiKeywordsText = string.Empty;

        public MainForm()
        {
            InitializeComponent();

            if (LicenseManager.UsageMode == LicenseUsageMode.Designtime || DesignMode)
            {
                return;
            }

            WireRuntimeEvents();
        }

        private void WireRuntimeEvents()
        {
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);
            this.cmbFolderPath.Leave += new EventHandler(this.HistoryComboBox_Leave);
            this.cmbFolderPath.KeyDown += new KeyEventHandler(this.HistoryComboBox_KeyDown);
            this.comboBox1.Leave += new EventHandler(this.HistoryComboBox_Leave);
            this.comboBox1.KeyDown += new KeyEventHandler(this.HistoryComboBox_KeyDown);
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            this.button1.Click += new System.EventHandler(this.btnGrep_Click);
            this.btnExportSakura.Click += new System.EventHandler(this.btnExportSakura_Click);
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            this.btnFileCopy.Click += new System.EventHandler(this.btnFileCopy_Click);
            this.btnMultiKeywords.Click += new System.EventHandler(this.btnMultiKeywords_Click);
            this.dataGridViewResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewResults_CellDoubleClick);
            this.FormClosing += new FormClosingEventHandler(MainForm_FormClosing);
            this.chkMethod.CheckedChanged += new System.EventHandler(this.chkMethod_CheckedChanged);
            this.cmbKeyword.Leave += new EventHandler(this.HistoryComboBox_Leave);
            this.cmbKeyword.KeyDown += new KeyEventHandler(this.cmbKeyword_KeyDown);
            this.cmbExcludeFolder.Leave += new EventHandler(this.HistoryComboBox_Leave);
            this.cmbExcludeFolder.KeyDown += new KeyEventHandler(this.HistoryComboBox_KeyDown);
            this.cmbExcludeExtension.Leave += new EventHandler(this.HistoryComboBox_Leave);
            this.cmbExcludeExtension.KeyDown += new KeyEventHandler(this.HistoryComboBox_KeyDown);
            this.txtFilePathFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtFileNameFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtRowNumFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtGrepResultFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtMethodFilter.TextChanged += new EventHandler(this.FilterTextChanged);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
            dataGridViewResults.Columns.Add("Encoding", "Encoding");
            dataGridViewResults.Columns["Encoding"].Visible = false;
            LoadSettings();
            UpdateMethodColumnVisibility();
            btnCancel.Enabled = false;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void dataGridViewResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                var row = dataGridViewResults.Rows[e.RowIndex];
                int filePathColumnIndex = dataGridViewResults.Columns["clmFilePath"].Index;
                int lineColumnIndex = dataGridViewResults.Columns["clmLine"].Index;

                string filePath = row.Cells[filePathColumnIndex].Value?.ToString() ?? string.Empty;

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (chkTagJump.Checked)
                {
                    string lineNumber = row.Cells[lineColumnIndex].Value?.ToString() ?? "1";
                    string sakuraPath = FindSakuraPath();

                    if (sakuraPath != null)
                    {
                        Process.Start(sakuraPath, $"-Y={lineNumber} \"{filePath}\"");
                    }
                    else
                    {
                        Process.Start(filePath);
                    }
                }
                else
                {
                    string directoryPath = Path.GetDirectoryName(filePath);
                    Process.Start("explorer.exe", $"\"{directoryPath}\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作を実行できませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string FindSakuraPath()
        {
            try
            {
                var configPath = ConfigurationManager.AppSettings["SakuraEditorPath"];
                if (!string.IsNullOrEmpty(configPath) && File.Exists(configPath))
                {
                    return configPath;
                }
            }
            catch (ConfigurationErrorsException)
            {
            }

            string[] searchPaths = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "sakura", "sakura.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "sakura", "sakura.exe")
            };

            foreach (var path in searchPaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }
            
            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (pathVar != null)
            {
                foreach (var p in pathVar.Split(Path.PathSeparator))
                {
                    var fullPath = Path.Combine(p, "sakura.exe");
                    if (File.Exists(fullPath))
                        return fullPath;
                }
            }

            return null;
        }


        private void LoadSettings()
        {
            if (!File.Exists(SettingsFileName)) return;

            try
            {
                using (var stream = new FileStream(SettingsFileName, FileMode.Open, FileAccess.Read))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    var settings = (AppSettings)serializer.ReadObject(stream);

                    LoadHistory(cmbFolderPath, settings.FolderPathHistory);
                    LoadHistory(comboBox1, settings.FilePatternHistory);
                    LoadHistory(cmbKeyword, settings.GrepPatternHistory);
                    LoadHistory(cmbExcludeFolder, MergeHistory(settings.ExcludeFolderHistory, settings.ExcludeFoldersText));
                    LoadHistory(cmbExcludeExtension, MergeHistory(settings.ExcludeExtensionHistory, settings.ExcludeExtensionsText));

                    chkSearchSubDir.Checked = settings.SearchSubDir;
                    chkCase.Checked = settings.CaseSensitive;
                    chkUseRegex.Checked = settings.UseRegex;
                    chkTagJump.Checked = settings.TagJump;
                    chkMethod.Checked = settings.DeriveMethod;
                    chkIgnoreComment.Checked = settings.IgnoreComment;
                    multiKeywordsText = settings.MultiKeywordsText ?? string.Empty;
                    chkIgnoreBinaryFile.Checked = settings.IgnoreBinaryFile;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                FolderPathHistory = GetHistory(cmbFolderPath),
                FilePatternHistory = GetHistory(comboBox1),
                GrepPatternHistory = GetHistory(cmbKeyword),
                ExcludeFolderHistory = GetHistory(cmbExcludeFolder),
                ExcludeExtensionHistory = GetHistory(cmbExcludeExtension),
                SearchSubDir = chkSearchSubDir.Checked,
                CaseSensitive = chkCase.Checked,
                UseRegex = chkUseRegex.Checked,
                TagJump = chkTagJump.Checked,
                DeriveMethod = chkMethod.Checked,
                IgnoreComment = chkIgnoreComment.Checked,
                MultiKeywordsText = multiKeywordsText,
                ExcludeFoldersText = cmbExcludeFolder.Text,
                ExcludeExtensionsText = cmbExcludeExtension.Text,
                IgnoreBinaryFile = chkIgnoreBinaryFile.Checked
            };

            try
            {
                using (var stream = new FileStream(SettingsFileName, FileMode.Create, FileAccess.Write))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    serializer.WriteObject(stream, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadHistory(ComboBox comboBox, List<string> history)
        {
            if (history != null && history.Any())
            {
                comboBox.Items.AddRange(history.ToArray());
                comboBox.Text = history.First();
            }
        }

        private List<string> GetHistory(ComboBox comboBox)
        {
            var history = new List<string>();
            if (!string.IsNullOrWhiteSpace(comboBox.Text))
            {
                history.Add(comboBox.Text);
            }
            history.AddRange(comboBox.Items.Cast<string>().Where(item => item != comboBox.Text));
            return history.Distinct().Take(MaxHistoryCount).ToList();
        }

        private static List<string> MergeHistory(List<string> history, string currentText)
        {
            var merged = new List<string>();
            if (!string.IsNullOrWhiteSpace(currentText))
            {
                merged.Add(currentText);
            }

            if (history != null)
            {
                merged.AddRange(history.Where(item => !string.IsNullOrWhiteSpace(item)));
            }

            return merged.Distinct().Take(MaxHistoryCount).ToList();
        }


        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    cmbFolderPath.Text = fbd.SelectedPath;
                }
            }
        }

        private async void btnGrep_Click(object sender, EventArgs e)
        {
            await RunSearchAsync(new[] { cmbKeyword.Text }, updateKeywordHistory: true);
        }

        private async Task RunSearchAsync(IReadOnlyList<string> grepPatterns, bool updateKeywordHistory)
        {
            string folderPath = cmbFolderPath.Text;
            string filePattern = comboBox1.Text;
            var normalizedPatterns = (grepPatterns ?? Array.Empty<string>())
                .Where(pattern => !string.IsNullOrWhiteSpace(pattern))
                .ToList();

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath) ||
                string.IsNullOrWhiteSpace(filePattern) || normalizedPatterns.Count == 0)
            {
                MessageBox.Show("検索フォルダー、ファイルパターン、検索パターンを正しく入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            UpdateHistory(cmbFolderPath, folderPath);
            UpdateHistory(comboBox1, filePattern);
            if (updateKeywordHistory && normalizedPatterns.Count > 0)
            {
                UpdateHistory(cmbKeyword, normalizedPatterns[0]);
            }

            dataGridViewResults.Rows.Clear();
            currentSearchResults.Clear();
            ClearAllFilters();
            button1.Enabled = false;
            btnMultiKeywords.Enabled = false;
            btnCancel.Enabled = true;
            this.Cursor = Cursors.WaitCursor;

            var stopwatch = Stopwatch.StartNew();
            labelTime.Text = "";

            bool wasCancelled = false;
            bool searchStarted = false;
            int totalFiles = 0;
            CancellationTokenSource localCancellationTokenSource = null;

            try
            {
                bool searchSubdirectories = chkSearchSubDir.Checked;
                bool caseSensitive = chkCase.Checked;
                bool useRegex = chkUseRegex.Checked;
                bool deriveMethod = chkMethod.Checked;
                bool ignoreComment = chkIgnoreComment.Checked;
                bool ignoreBinaryFile = chkIgnoreBinaryFile.Checked;
                var excludeFolderPatterns = ParseExcludeFolderPatterns(cmbExcludeFolder.Text);
                var excludeExtensions = ParseExcludeExtensions(cmbExcludeExtension.Text);
                if (useRegex && excludeFolderPatterns.Count > 0)
                {
                    string invalidExcludePattern = GetInvalidRegexPattern(excludeFolderPatterns);
                    if (invalidExcludePattern != null)
                    {
                        MessageBox.Show($"除外フォルダの正規表現が不正です: {invalidExcludePattern}", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                var searchOption = searchSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                UpdateHistory(cmbExcludeFolder, cmbExcludeFolder.Text);
                UpdateHistory(cmbExcludeExtension, cmbExcludeExtension.Text);
                string[] filesToSearch;
                try
                {
                    filesToSearch = Directory.GetFiles(folderPath, filePattern, searchOption);
                    if (excludeFolderPatterns.Count > 0)
                    {
                        filesToSearch = filesToSearch
                            .Where(filePath => !ContainsExcludedFolder(filePath, excludeFolderPatterns, caseSensitive, useRegex))
                            .ToArray();
                    }
                    if (excludeExtensions.Count > 0)
                    {
                        filesToSearch = filesToSearch
                            .Where(filePath => !HasExcludedExtension(filePath, excludeExtensions, caseSensitive))
                            .ToArray();
                    }
                    if (ignoreBinaryFile)
                    {
                        filesToSearch = filesToSearch
                            .Where(filePath => !LooksLikeBinaryFile(filePath))
                            .ToArray();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイル一覧の取得に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                totalFiles = filesToSearch.Length;

                if (totalFiles == 0)
                {
                    MessageBox.Show("対象ファイルが見つかりません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressBar.Value = 0;
                    lblPer.Text = "0 %";
                    return;
                }

                progressBar.Maximum = totalFiles;
                progressBar.Value = 0;
                lblPer.Text = "0 %";

                searchCancellationTokenSource?.Cancel();
                searchCancellationTokenSource?.Dispose();

                localCancellationTokenSource = new CancellationTokenSource();
                searchCancellationTokenSource = localCancellationTokenSource;
                var cancellationToken = localCancellationTokenSource.Token;

                var progress = new Progress<int>(processedCount =>
                {
                    int clamped = Math.Min(processedCount, totalFiles);
                    progressBar.Value = clamped;
                    int percentage = totalFiles == 0 ? 0 : (int)((double)clamped / totalFiles * 100);
                    lblPer.Text = $"{percentage} %";
                });

                List<SearchResult> searchResults = null;
                searchStarted = true;

                Action<Exception> errorHandler = ex =>
                {
                    try
                    {
                        BeginInvoke((Action)(() =>
                        {
                            MessageBox.Show($"ディレクトリの検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                    catch (ObjectDisposedException)
                    {
                        // フォームが既に破棄されている場合は無視
                    }
                };

                searchResults = await Task.Run(() =>
                    GrepEngine.SearchFiles(
                        filesToSearch,
                        normalizedPatterns,
                        caseSensitive,
                        useRegex,
                        deriveMethod,
                        ignoreComment,
                        progress,
                        cancellationToken,
                        errorHandler).ToList(),
                    cancellationToken);
                if (localCancellationTokenSource?.IsCancellationRequested == true)
                {
                    wasCancelled = true;
                }

                if (!wasCancelled && searchResults != null)
                {
                    currentSearchResults = searchResults ?? new List<SearchResult>();
                    ApplyFilters();
                }
            }
            catch (OperationCanceledException)
            {
                wasCancelled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                stopwatch.Stop();

                if (searchStarted)
                {
                    labelTime.Text = wasCancelled ? $"Time: {stopwatch.Elapsed:mm\\:ss} (中止)" : $"Time: {stopwatch.Elapsed:mm\\:ss}";

                    if (!wasCancelled)
                    {
                        progressBar.Value = totalFiles;
                        lblPer.Text = "100 %";
                    }
                    else
                    {
                        int percentage = totalFiles == 0 ? 0 : (int)((double)progressBar.Value / totalFiles * 100);
                        lblPer.Text = $"{percentage} % (中止)";
                    }
                }

                button1.Enabled = true;
                btnMultiKeywords.Enabled = true;
                btnCancel.Enabled = false;
                this.Cursor = Cursors.Default;

                localCancellationTokenSource?.Dispose();
                searchCancellationTokenSource = null;
            }
        }

        private async void btnMultiKeywords_Click(object sender, EventArgs e)
        {
            using (var dialog = new MultiKeywordsForm(multiKeywordsText))
            {
                var dialogResult = dialog.ShowDialog(this);
                multiKeywordsText = dialog.MultiKeywordsText;
                if (dialogResult == DialogResult.OK)
                {
                    var keywords = dialog.Keywords?.ToList() ?? new List<string>();
                    await RunSearchAsync(keywords, updateKeywordHistory: false);
                }
            }
        }
        
        private void UpdateHistory(ComboBox comboBox, string newItem)
        {
            if (comboBox == null || string.IsNullOrWhiteSpace(newItem))
            {
                return;
            }

            var items = comboBox.Items.Cast<string>().ToList();
            items.Remove(newItem);
            items.Insert(0, newItem);
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.Take(MaxHistoryCount).ToArray());
            comboBox.Text = newItem;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (searchCancellationTokenSource != null && !searchCancellationTokenSource.IsCancellationRequested)
            {
                searchCancellationTokenSource.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private void cmbKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                UpdateComboBoxHistory(sender);
                if (button1.Enabled)
                {
                    button1.PerformClick();
                }
            }
        }

        private void HistoryComboBox_Leave(object sender, EventArgs e)
        {
            UpdateComboBoxHistory(sender);
        }

        private void HistoryComboBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                UpdateComboBoxHistory(sender);
            }
        }

        private void UpdateComboBoxHistory(object sender)
        {
            var comboBox = sender as ComboBox;
            if (comboBox == null)
            {
                return;
            }

            UpdateHistory(comboBox, comboBox.Text);
        }

        private void FilterTextChanged(object sender, EventArgs e)
        {
            ApplyFilters();
        }

        private void ApplyFilters()
        {
            var filtered = currentSearchResults?
                .Where(MatchesFilters)
                .ToList() ?? new List<SearchResult>();

            RenderResults(filtered);
        }

        private void ClearAllFilters()
        {
            txtFilePathFilter.Text = string.Empty;
            txtFileNameFilter.Text = string.Empty;
            txtRowNumFilter.Text = string.Empty;
            txtGrepResultFilter.Text = string.Empty;
            txtMethodFilter.Text = string.Empty;
        }

        private bool MatchesFilters(SearchResult result)
        {
            if (result == null)
            {
                return false;
            }

            string filePathFilter = GetFilterText(txtFilePathFilter);
            if (!string.IsNullOrEmpty(filePathFilter) && !ContainsText(result.FilePath, filePathFilter))
            {
                return false;
            }

            string fileNameFilter = GetFilterText(txtFileNameFilter);
            if (!string.IsNullOrEmpty(fileNameFilter) && !ContainsText(result.FileName, fileNameFilter))
            {
                return false;
            }

            string rowFilter = GetFilterText(txtRowNumFilter);
            if (!string.IsNullOrEmpty(rowFilter) && !ContainsText(result.LineNumber.ToString(), rowFilter))
            {
                return false;
            }

            string grepResultFilter = GetFilterText(txtGrepResultFilter);
            if (!string.IsNullOrEmpty(grepResultFilter) && !ContainsText(result.LineText, grepResultFilter))
            {
                return false;
            }

            string methodFilter = GetFilterText(txtMethodFilter);
            if (!string.IsNullOrEmpty(methodFilter) && !ContainsText(result.MethodSignature, methodFilter))
            {
                return false;
            }

            return true;
        }

        private static string GetFilterText(TextBox textBox)
        {
            return textBox?.Text?.Trim() ?? string.Empty;
        }

        private static bool ContainsText(string source, string filter)
        {
            if (string.IsNullOrEmpty(filter))
            {
                return true;
            }

            if (string.IsNullOrEmpty(source))
            {
                return false;
            }

            return source.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void RenderResults(IReadOnlyList<SearchResult> results)
        {
            dataGridViewResults.SuspendLayout();
            dataGridViewResults.Rows.Clear();

            if (results != null && results.Count > 0)
            {
                var rows = new DataGridViewRow[results.Count];
                for (int i = 0; i < results.Count; i++)
                {
                    var result = results[i];
                    object[] cells =
                    {
                        result.FilePath,
                        result.FileName,
                        result.LineNumber,
                        result.LineText,
                        result.MethodSignature,
                        result.EncodingName
                    };

                    var row = new DataGridViewRow();
                    row.CreateCells(dataGridViewResults, cells);
                    rows[i] = row;
                }

                dataGridViewResults.Rows.AddRange(rows);
            }

            dataGridViewResults.ResumeLayout();
        }

        private void btnFileCopy_Click(object sender, EventArgs e)
        {
            var selectedIndices = dataGridViewResults.SelectedCells.Cast<DataGridViewCell>()
                .Select(cell => cell.RowIndex)
                .Concat(dataGridViewResults.SelectedRows.Cast<DataGridViewRow>().Select(row => row.Index))
                .Distinct()
                .ToList();

            if (selectedIndices.Count == 0)
            {
                MessageBox.Show("ファイルをコピーする行を選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var files = new StringCollection();
            int filePathColumnIndex = dataGridViewResults.Columns["clmFilePath"].Index;
            foreach (var index in selectedIndices)
            {
                if (index < 0 || index >= dataGridViewResults.Rows.Count)
                {
                    continue;
                }

                var cellValue = dataGridViewResults.Rows[index].Cells[filePathColumnIndex].Value;
                if (cellValue == null)
                {
                    continue;
                }

                var filePath = cellValue.ToString();
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    files.Add(filePath);
                }
            }

            if (files.Count == 0)
            {
                MessageBox.Show("実在するファイルが選択されていません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Clipboard.SetFileDropList(files);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkMethod_CheckedChanged(object sender, EventArgs e)
        {
            UpdateMethodColumnVisibility();
        }

        private void UpdateMethodColumnVisibility()
        {
            if (clmMethodSignature != null)
            {
                clmMethodSignature.Visible = chkMethod.Checked;
            }
        }

        // ★★ここから修正★★
        private void btnExportSakura_Click(object sender, EventArgs e)
        {
            if (dataGridViewResults.Rows.Count == 0)
            {
                MessageBox.Show("エクスポートするデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string fileName = Path.Combine(AppContext.BaseDirectory, DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".grep");
            try
            {
                using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
                {
                    int filePathColumnIndex = dataGridViewResults.Columns["clmFilePath"].Index;
                    int lineColumnIndex = dataGridViewResults.Columns["clmLine"].Index;
                    int resultColumnIndex = dataGridViewResults.Columns["clmGrepResult"].Index;

                    foreach (DataGridViewRow row in dataGridViewResults.Rows)
                    {
                        string filePath = row.Cells[filePathColumnIndex].Value?.ToString() ?? string.Empty;
                        string lineNumber = row.Cells[lineColumnIndex].Value?.ToString() ?? string.Empty;
                        string lineContent = row.Cells[resultColumnIndex].Value?.ToString() ?? string.Empty;
                        var encodingCell = row.Cells["Encoding"];
                        string encoding = encodingCell != null && encodingCell.Value != null ? encodingCell.Value.ToString() : "UTF-8";
                        
                        writer.WriteLine($"{filePath}({lineNumber},1)  [{encoding}]: {lineContent}");
                    }
                }
                
                string sakuraPath = FindSakuraPath();
                if (sakuraPath != null)
                {
                    try
                    {
                        Process.Start(sakuraPath, $"\"{fileName}\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルのエクスポートには成功しましたが、サクラエディタの起動に失敗しました。\nエラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"{fileName} に結果を保存しました。\n(サクラエディタが見つからなかったため、ファイルは自動で開かれませんでした)", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // ★★ここまで修正★★

        private void cmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void cmbFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length > 0)
            {
                string path = files[0];
                if (Directory.Exists(path))
                {
                    cmbFolderPath.Text = path;
                }
                else if (File.Exists(path))
                {
                    cmbFolderPath.Text = Path.GetDirectoryName(path);
                }
            }
        }

        private static IReadOnlyList<string> ParseExcludeFolderPatterns(string excludeFoldersText)
        {
            if (string.IsNullOrWhiteSpace(excludeFoldersText))
            {
                return Array.Empty<string>();
            }

            return excludeFoldersText
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(pattern => pattern.Trim())
                .Where(pattern => !string.IsNullOrEmpty(pattern))
                .ToArray();
        }

        private static bool ContainsExcludedFolder(string filePath, IReadOnlyList<string> excludeFolderPatterns, bool caseSensitive, bool useRegex)
        {
            if (string.IsNullOrWhiteSpace(filePath) || excludeFolderPatterns == null || excludeFolderPatterns.Count == 0)
            {
                return false;
            }

            string directoryPath = Path.GetDirectoryName(filePath);
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                return false;
            }

            var folderNames = directoryPath
                .Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)
                .Select(Path.GetFileName)
                .Where(folderName => !string.IsNullOrEmpty(folderName))
                .ToArray();

            if (useRegex)
            {
                var regexOptions = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
                foreach (string pattern in excludeFolderPatterns)
                {
                    var regex = new Regex(pattern, regexOptions);
                    if (folderNames.Any(folderName => regex.IsMatch(folderName)))
                    {
                        return true;
                    }
                }

                return false;
            }

            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            return folderNames.Any(folderName => excludeFolderPatterns.Contains(folderName, comparer));
        }

        private static IReadOnlyList<string> ParseExcludeExtensions(string excludeExtensionsText)
        {
            if (string.IsNullOrWhiteSpace(excludeExtensionsText))
            {
                return Array.Empty<string>();
            }

            return excludeExtensionsText
                .Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(extension => extension.Trim())
                .Where(extension => !string.IsNullOrEmpty(extension))
                .Select(extension => extension.StartsWith(".") ? extension : "." + extension)
                .ToArray();
        }

        private static bool HasExcludedExtension(string filePath, IReadOnlyList<string> excludeExtensions, bool caseSensitive)
        {
            if (string.IsNullOrWhiteSpace(filePath) || excludeExtensions == null || excludeExtensions.Count == 0)
            {
                return false;
            }

            string extension = Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(extension))
            {
                return false;
            }

            var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            return excludeExtensions.Contains(extension, comparer);
        }

        private static bool LooksLikeBinaryFile(string filePath)
        {
            const int BufferSize = 4096;

            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize, FileOptions.SequentialScan))
                {
                    var buffer = new byte[Math.Min(BufferSize, Math.Max(1, (int)Math.Min(stream.Length, BufferSize)))];
                    int bytesRead = stream.Read(buffer, 0, buffer.Length);
                    for (int i = 0; i < bytesRead; i++)
                    {
                        if (buffer[i] == 0)
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private static string GetInvalidRegexPattern(IEnumerable<string> patterns)
        {
            foreach (string pattern in patterns)
            {
                try
                {
                    new Regex(pattern);
                }
                catch (ArgumentException)
                {
                    return pattern;
                }
            }

            return null;
        }

        [DataContract]
        private class AppSettings
        {
            [DataMember]
            public List<string> FolderPathHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> FilePatternHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> GrepPatternHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> ExcludeFolderHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> ExcludeExtensionHistory { get; set; } = new List<string>();
            [DataMember]
            public bool SearchSubDir { get; set; }
            [DataMember]
            public bool CaseSensitive { get; set; }
            [DataMember]
            public bool UseRegex { get; set; }
            
            [DataMember]
            public bool TagJump { get; set; } = true;

            [DataMember]
            public bool DeriveMethod { get; set; }

            [DataMember]
            public bool IgnoreComment { get; set; }

            [DataMember]
            public string MultiKeywordsText { get; set; } = string.Empty;

            [DataMember]
            public string ExcludeFoldersText { get; set; } = string.Empty;

            [DataMember]
            public string ExcludeExtensionsText { get; set; } = string.Empty;

            [DataMember]
            public bool IgnoreBinaryFile { get; set; }
        }
    }
}
