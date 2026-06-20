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
        private const string SakuraSearchConditionsBegin = "# SimpleGrep Search Conditions Begin";
        private const string SakuraSearchConditionsEnd = "# SimpleGrep Search Conditions End";
        private const int MaxHistoryCount = 10;
        private CancellationTokenSource searchCancellationTokenSource;
        private CancellationTokenSource filterCancellationTokenSource;
        private List<SearchResult> currentSearchResults = new List<SearchResult>();
        private List<SearchResult> displayedSearchResults = new List<SearchResult>();
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
            this.btnExportSettings.Click += new System.EventHandler(this.btnExportSettings_Click);
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            this.btnFileCopy.Click += new System.EventHandler(this.btnFileCopy_Click);
            this.btnMultiKeywords.Click += new System.EventHandler(this.btnMultiKeywords_Click);
            this.btnCollectExtensions.Click += new System.EventHandler(this.btnCollectExtensions_Click);
            this.btnClearFilters.Click += new System.EventHandler(this.btnClearFilters_Click);
            this.dataGridViewResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewResults_CellDoubleClick);
            InitializeResultsContextMenu();
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
            this.txtExtensionFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtRowNumFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtGrepResultFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.txtMethodFilter.TextChanged += new EventHandler(this.FilterTextChanged);
            this.dataGridViewResults.DragEnter += new DragEventHandler(this.dataGridViewResults_DragEnter);
            this.dataGridViewResults.DragDrop += new DragEventHandler(this.dataGridViewResults_DragDrop);
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

        private void InitializeResultsContextMenu()
        {
            var contextMenu = new ContextMenuStrip();
            var copyAllMenuItem = new ToolStripMenuItem("全てコピー");
            copyAllMenuItem.Click += new EventHandler(this.CopyAllResultsMenuItem_Click);
            contextMenu.Items.Add(copyAllMenuItem);
            dataGridViewResults.ContextMenuStrip = contextMenu;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            filterCancellationTokenSource?.Cancel();
            filterCancellationTokenSource?.Dispose();
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
            try
            {
                WriteSettingsFile(SettingsFileName, GetCurrentSettings());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private AppSettings GetCurrentSettings()
        {
            return new AppSettings
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
        }

        private static void WriteSettingsFile(string fileName, AppSettings settings)
        {
            using (var stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                serializer.WriteObject(stream, settings);
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
            var filePatterns = ParseFilePatterns(filePattern);
            var normalizedPatterns = (grepPatterns ?? Array.Empty<string>())
                .Where(pattern => !string.IsNullOrWhiteSpace(pattern))
                .ToList();

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath) ||
                filePatterns.Count == 0 || normalizedPatterns.Count == 0)
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
            lblResultCount.Text = string.Empty;
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

                searchCancellationTokenSource?.Cancel();
                searchCancellationTokenSource?.Dispose();

                localCancellationTokenSource = new CancellationTokenSource();
                searchCancellationTokenSource = localCancellationTokenSource;
                var cancellationToken = localCancellationTokenSource.Token;

                searchStarted = true;
                string[] filesToSearch;
                try
                {
                    filesToSearch = await Task.Run(() =>
                        GetFilesToSearch(
                            folderPath,
                            filePatterns,
                            searchOption,
                            excludeFolderPatterns,
                            excludeExtensions,
                            caseSensitive,
                            useRegex,
                            ignoreBinaryFile,
                            cancellationToken),
                        cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    wasCancelled = true;
                    return;
                }
                catch (Exception ex)
                {
                    searchStarted = false;
                    MessageBox.Show($"ファイル一覧の取得に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                totalFiles = filesToSearch.Length;

                if (totalFiles == 0)
                {
                    MessageBox.Show("対象ファイルが見つかりません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressBar.Value = 0;
                    lblPer.Text = "0 %";
                    searchStarted = false;
                    return;
                }

                progressBar.Maximum = totalFiles;
                progressBar.Value = 0;
                lblPer.Text = "0 %";

                var progress = new Progress<int>(processedCount =>
                {
                    int clamped = Math.Min(processedCount, totalFiles);
                    progressBar.Value = clamped;
                    int percentage = totalFiles == 0 ? 0 : (int)((double)clamped / totalFiles * 100);
                    lblPer.Text = $"{percentage} %";
                });

                List<SearchResult> searchResults = null;

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

                if (searchResults != null)
                {
                    currentSearchResults = SortSearchResults(searchResults);
                    await ApplyFiltersAsync();
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

        private void btnCollectExtensions_Click(object sender, EventArgs e)
        {
            var form = new ExtensionCollectorForm(cmbFolderPath.Text);
            form.Show(this);
        }

        private void btnClearFilters_Click(object sender, EventArgs e)
        {
            ClearAllFilters();
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

        private async void FilterTextChanged(object sender, EventArgs e)
        {
            await ApplyFiltersAsync();
        }

        private async Task ApplyFiltersAsync()
        {
            filterCancellationTokenSource?.Cancel();
            filterCancellationTokenSource?.Dispose();
            filterCancellationTokenSource = new CancellationTokenSource();
            var cancellationToken = filterCancellationTokenSource.Token;

            var source = currentSearchResults?.ToList() ?? new List<SearchResult>();
            var criteria = GetFilterCriteria();
            List<SearchResult> filtered;

            try
            {
                filtered = await Task.Run(() => source
                    .Where(result => MatchesFilters(result, criteria, cancellationToken))
                    .ToList(), cancellationToken);
            }
            catch (OperationCanceledException)
            {
                return;
            }

            if (cancellationToken.IsCancellationRequested)
            {
                return;
            }

            RenderResults(filtered);
        }

        private static List<SearchResult> SortSearchResults(IEnumerable<SearchResult> results)
        {
            return (results ?? Enumerable.Empty<SearchResult>())
                .OrderBy(result => result.FilePath ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(result => result.LineNumber)
                .ThenBy(result => result.LineText ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private void ClearAllFilters()
        {
            txtFilePathFilter.Text = string.Empty;
            txtFileNameFilter.Text = string.Empty;
            txtExtensionFilter.Text = string.Empty;
            txtRowNumFilter.Text = string.Empty;
            txtGrepResultFilter.Text = string.Empty;
            txtMethodFilter.Text = string.Empty;
        }

        private FilterCriteria GetFilterCriteria()
        {
            return new FilterCriteria
            {
                FilePath = GetFilterText(txtFilePathFilter),
                FileName = GetFilterText(txtFileNameFilter),
                Extension = GetFilterText(txtExtensionFilter),
                RowNumber = GetFilterText(txtRowNumFilter),
                GrepResult = GetFilterText(txtGrepResultFilter),
                Method = GetFilterText(txtMethodFilter)
            };
        }

        private static bool MatchesFilters(SearchResult result, FilterCriteria criteria, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (result == null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.FilePath) && !ContainsText(result.FilePath, criteria.FilePath))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.FileName) && !ContainsText(result.FileName, criteria.FileName))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.Extension) && !ContainsText(result.FileExtension, criteria.Extension))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.RowNumber) && !ContainsText(result.LineNumber.ToString(), criteria.RowNumber))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.GrepResult) && !ContainsText(result.LineText, criteria.GrepResult))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(criteria.Method) && !ContainsText(result.MethodSignature, criteria.Method))
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
            displayedSearchResults = results?.ToList() ?? new List<SearchResult>();
            lblResultCount.Text = displayedSearchResults.Count > 0 ? $"{displayedSearchResults.Count} 件" : string.Empty;
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
                        result.FileExtension,
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

        private async void CopyAllResultsMenuItem_Click(object sender, EventArgs e)
        {
            var rows = displayedSearchResults?.ToList() ?? new List<SearchResult>();
            if (rows.Count == 0)
            {
                MessageBox.Show("コピーするデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                var visibleColumns = dataGridViewResults.Columns
                    .Cast<DataGridViewColumn>()
                    .Where(column => column.Visible)
                    .OrderBy(column => column.DisplayIndex)
                    .Select(column => new ResultColumnInfo(column.Name, column.HeaderText))
                    .ToList();

                string clipboardText = await Task.Run(() => BuildResultsClipboardText(rows, visibleColumns));
                Clipboard.SetText(clipboardText);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string BuildResultsClipboardText(IReadOnlyList<SearchResult> rows, IReadOnlyList<ResultColumnInfo> visibleColumns)
        {
            var builder = new StringBuilder();
            builder.AppendLine(string.Join("\t", visibleColumns.Select(column => NormalizeClipboardCell(column.HeaderText))));

            foreach (var row in rows)
            {
                var values = visibleColumns.Select(column => NormalizeClipboardCell(GetResultColumnValue(row, column.Name)));
                builder.AppendLine(string.Join("\t", values));
            }

            return builder.ToString();
        }

        private static string GetResultColumnValue(SearchResult result, string columnName)
        {
            if (result == null)
            {
                return string.Empty;
            }

            switch (columnName)
            {
                case "clmFilePath":
                    return result.FilePath ?? string.Empty;
                case "clmFileName":
                    return result.FileName ?? string.Empty;
                case "clmExtension":
                    return result.FileExtension ?? string.Empty;
                case "clmLine":
                    return result.LineNumber.ToString();
                case "clmGrepResult":
                    return result.LineText ?? string.Empty;
                case "clmMethodSignature":
                    return result.MethodSignature ?? string.Empty;
                case "Encoding":
                    return result.EncodingName ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        private static string NormalizeClipboardCell(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("\r\n", " ")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\t", " ");
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
        private async void btnExportSakura_Click(object sender, EventArgs e)
        {
            var results = displayedSearchResults?.ToList() ?? new List<SearchResult>();
            if (results.Count == 0)
            {
                MessageBox.Show("エクスポートするデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string fileName = Path.Combine(AppContext.BaseDirectory, DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".grep");
            try
            {
                var searchConditions = GetCurrentSakuraSearchConditions();
                await Task.Run(() => WriteSakuraGrepFile(fileName, results, searchConditions));
                
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

        private void btnExportSettings_Click(object sender, EventArgs e)
        {
            string fileName = Path.Combine(
                AppContext.BaseDirectory,
                "SimpleGrep.settings_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".json");

            try
            {
                WriteSettingsFile(fileName, GetCurrentSettings());
                MessageBox.Show($"{fileName} に設定を保存しました。", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定のエクスポートに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private SakuraSearchConditions GetCurrentSakuraSearchConditions()
        {
            return new SakuraSearchConditions
            {
                ExportedAt = DateTime.Now,
                SearchFolder = cmbFolderPath.Text,
                FilePattern = comboBox1.Text,
                Keyword = cmbKeyword.Text,
                MultiKeywords = multiKeywordsText,
                ExcludeFolders = cmbExcludeFolder.Text,
                ExcludeExtensions = cmbExcludeExtension.Text,
                SearchSubDir = chkSearchSubDir.Checked,
                CaseSensitive = chkCase.Checked,
                UseRegex = chkUseRegex.Checked,
                DeriveMethod = chkMethod.Checked,
                IgnoreComment = chkIgnoreComment.Checked,
                IgnoreBinaryFile = chkIgnoreBinaryFile.Checked
            };
        }

        private static void WriteSakuraGrepFile(string fileName, IReadOnlyList<SearchResult> results, SakuraSearchConditions searchConditions)
        {
            using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
            {
                WriteSakuraSearchConditions(writer, searchConditions);

                foreach (var result in results)
                {
                    string filePath = result.FilePath ?? string.Empty;
                    string lineNumber = result.LineNumber.ToString();
                    string lineContent = result.LineText ?? string.Empty;
                    string encoding = string.IsNullOrEmpty(result.EncodingName) ? "UTF-8" : result.EncodingName;

                    writer.WriteLine($"{filePath}({lineNumber},1)  [{encoding}]: {lineContent}");
                }
            }
        }

        private static void WriteSakuraSearchConditions(StreamWriter writer, SakuraSearchConditions searchConditions)
        {
            if (searchConditions == null)
            {
                searchConditions = new SakuraSearchConditions();
            }

            writer.WriteLine(SakuraSearchConditionsBegin);
            writer.WriteLine("# ExportedAt: " + searchConditions.ExportedAt.ToString("yyyy-MM-dd HH:mm:ss"));
            writer.WriteLine("# SearchFolder: " + EscapeSearchConditionValue(searchConditions.SearchFolder));
            writer.WriteLine("# FilePattern: " + EscapeSearchConditionValue(searchConditions.FilePattern));
            writer.WriteLine("# Keyword: " + EscapeSearchConditionValue(searchConditions.Keyword));
            writer.WriteLine("# MultiKeywords: " + EscapeSearchConditionValue(searchConditions.MultiKeywords));
            writer.WriteLine("# ExcludeFolders: " + EscapeSearchConditionValue(searchConditions.ExcludeFolders));
            writer.WriteLine("# ExcludeExtensions: " + EscapeSearchConditionValue(searchConditions.ExcludeExtensions));
            writer.WriteLine("# SearchSubDir: " + searchConditions.SearchSubDir);
            writer.WriteLine("# CaseSensitive: " + searchConditions.CaseSensitive);
            writer.WriteLine("# UseRegex: " + searchConditions.UseRegex);
            writer.WriteLine("# DeriveMethod: " + searchConditions.DeriveMethod);
            writer.WriteLine("# IgnoreComment: " + searchConditions.IgnoreComment);
            writer.WriteLine("# IgnoreBinaryFile: " + searchConditions.IgnoreBinaryFile);
            writer.WriteLine(SakuraSearchConditionsEnd);
        }

        private void dataGridViewResults_DragEnter(object sender, DragEventArgs e)
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

        private void dataGridViewResults_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files == null || files.Length == 0)
            {
                return;
            }

            string fileName = files[0];
            var dialogResult = MessageBox.Show(
                $"{fileName}\n\n検索結果と検索条件をインポートしますか？",
                "インポート確認",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (dialogResult != DialogResult.OK)
            {
                return;
            }

            try
            {
                var importData = ReadSakuraGrepFile(fileName);
                if (importData.Results.Count == 0)
                {
                    MessageBox.Show("復元できる検索結果が見つかりませんでした。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (importData.HasSearchConditions)
                {
                    ApplySakuraSearchConditions(importData.SearchConditions);
                }

                currentSearchResults = importData.Results;
                ClearAllFilters();
                RenderResults(currentSearchResults);
                labelTime.Text = "Restored from file";
                progressBar.Value = 0;
                lblPer.Text = "0 %";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"検索結果の復元に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplySakuraSearchConditions(SakuraSearchConditions searchConditions)
        {
            if (searchConditions == null)
            {
                return;
            }

            cmbFolderPath.Text = searchConditions.SearchFolder ?? string.Empty;
            comboBox1.Text = searchConditions.FilePattern ?? string.Empty;
            cmbKeyword.Text = searchConditions.Keyword ?? string.Empty;
            multiKeywordsText = searchConditions.MultiKeywords ?? string.Empty;
            cmbExcludeFolder.Text = searchConditions.ExcludeFolders ?? string.Empty;
            cmbExcludeExtension.Text = searchConditions.ExcludeExtensions ?? string.Empty;
            chkSearchSubDir.Checked = searchConditions.SearchSubDir;
            chkCase.Checked = searchConditions.CaseSensitive;
            chkUseRegex.Checked = searchConditions.UseRegex;
            chkMethod.Checked = searchConditions.DeriveMethod;
            chkIgnoreComment.Checked = searchConditions.IgnoreComment;
            chkIgnoreBinaryFile.Checked = searchConditions.IgnoreBinaryFile;
            UpdateMethodColumnVisibility();
        }

        private static SakuraGrepImportData ReadSakuraGrepFile(string fileName)
        {
            var results = new List<SearchResult>();
            var searchConditions = new SakuraSearchConditions();
            bool hasSearchConditions = false;
            bool skippingSearchConditions = false;

            foreach (string line in File.ReadLines(fileName, Encoding.UTF8))
            {
                if (line == SakuraSearchConditionsBegin)
                {
                    hasSearchConditions = true;
                    skippingSearchConditions = true;
                    continue;
                }

                if (line == SakuraSearchConditionsEnd)
                {
                    skippingSearchConditions = false;
                    continue;
                }

                if (skippingSearchConditions || string.IsNullOrWhiteSpace(line))
                {
                    if (skippingSearchConditions)
                    {
                        TryApplySakuraSearchConditionLine(line, searchConditions);
                    }
                    continue;
                }

                SearchResult result;
                if (TryParseSakuraGrepLine(line, out result))
                {
                    results.Add(result);
                }
            }

            return new SakuraGrepImportData
            {
                Results = results,
                SearchConditions = searchConditions,
                HasSearchConditions = hasSearchConditions
            };
        }

        private static void TryApplySakuraSearchConditionLine(string line, SakuraSearchConditions searchConditions)
        {
            if (string.IsNullOrEmpty(line) || searchConditions == null || !line.StartsWith("# "))
            {
                return;
            }

            int separatorIndex = line.IndexOf(": ", StringComparison.Ordinal);
            if (separatorIndex < 0)
            {
                return;
            }

            string key = line.Substring(2, separatorIndex - 2);
            string value = line.Substring(separatorIndex + 2);
            switch (key)
            {
                case "ExportedAt":
                    DateTime exportedAt;
                    if (DateTime.TryParse(value, out exportedAt))
                    {
                        searchConditions.ExportedAt = exportedAt;
                    }
                    break;
                case "SearchFolder":
                    searchConditions.SearchFolder = UnescapeSearchConditionValue(value);
                    break;
                case "FilePattern":
                    searchConditions.FilePattern = UnescapeSearchConditionValue(value);
                    break;
                case "Keyword":
                    searchConditions.Keyword = UnescapeSearchConditionValue(value);
                    break;
                case "MultiKeywords":
                    searchConditions.MultiKeywords = UnescapeSearchConditionValue(value);
                    break;
                case "ExcludeFolders":
                    searchConditions.ExcludeFolders = UnescapeSearchConditionValue(value);
                    break;
                case "ExcludeExtensions":
                    searchConditions.ExcludeExtensions = UnescapeSearchConditionValue(value);
                    break;
                case "SearchSubDir":
                    searchConditions.SearchSubDir = ParseBooleanSearchCondition(value);
                    break;
                case "CaseSensitive":
                    searchConditions.CaseSensitive = ParseBooleanSearchCondition(value);
                    break;
                case "UseRegex":
                    searchConditions.UseRegex = ParseBooleanSearchCondition(value);
                    break;
                case "DeriveMethod":
                    searchConditions.DeriveMethod = ParseBooleanSearchCondition(value);
                    break;
                case "IgnoreComment":
                    searchConditions.IgnoreComment = ParseBooleanSearchCondition(value);
                    break;
                case "IgnoreBinaryFile":
                    searchConditions.IgnoreBinaryFile = ParseBooleanSearchCondition(value);
                    break;
            }
        }

        private static bool ParseBooleanSearchCondition(string value)
        {
            bool result;
            return bool.TryParse(value, out result) && result;
        }

        private static string EscapeSearchConditionValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return "base64:" + Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private static string UnescapeSearchConditionValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            const string Base64Prefix = "base64:";
            if (!value.StartsWith(Base64Prefix, StringComparison.Ordinal))
            {
                return value;
            }

            try
            {
                byte[] bytes = Convert.FromBase64String(value.Substring(Base64Prefix.Length));
                return Encoding.UTF8.GetString(bytes);
            }
            catch (FormatException)
            {
                return value;
            }
        }

        private static bool TryParseSakuraGrepLine(string line, out SearchResult result)
        {
            result = null;
            var match = Regex.Match(line ?? string.Empty, @"^(?<path>.+)\((?<line>\d+),\d+\)\s+\[(?<encoding>[^\]]*)\]:(?<text>.*)$");
            if (!match.Success)
            {
                return false;
            }

            int lineNumber;
            if (!int.TryParse(match.Groups["line"].Value, out lineNumber))
            {
                return false;
            }

            string filePath = match.Groups["path"].Value;
            string lineText = match.Groups["text"].Value;
            if (lineText.StartsWith(" "))
            {
                lineText = lineText.Substring(1);
            }

            result = new SearchResult
            {
                FilePath = filePath,
                FileName = Path.GetFileName(filePath),
                FileExtension = Path.GetExtension(filePath),
                LineNumber = lineNumber,
                LineText = lineText,
                MethodSignature = string.Empty,
                EncodingName = match.Groups["encoding"].Value
            };
            return true;
        }

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

        private static string[] GetFilesToSearch(
            string folderPath,
            IReadOnlyList<string> filePatterns,
            SearchOption searchOption,
            IReadOnlyList<string> excludeFolderPatterns,
            IReadOnlyList<string> excludeExtensions,
            bool caseSensitive,
            bool useRegex,
            bool ignoreBinaryFile,
            CancellationToken cancellationToken)
        {
            var files = filePatterns
                .SelectMany(filePattern =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return Directory.EnumerateFiles(folderPath, filePattern, searchOption);
                })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return true;
                });

            if (excludeFolderPatterns != null && excludeFolderPatterns.Count > 0)
            {
                files = files.Where(filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return !ContainsExcludedFolder(filePath, excludeFolderPatterns, caseSensitive, useRegex);
                });
            }

            if (excludeExtensions != null && excludeExtensions.Count > 0)
            {
                files = files.Where(filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return !HasExcludedExtension(filePath, excludeExtensions, caseSensitive);
                });
            }

            if (ignoreBinaryFile)
            {
                files = files.Where(filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    return !LooksLikeBinaryFile(filePath);
                });
            }

            return files.ToArray();
        }

        private static IReadOnlyList<string> ParseFilePatterns(string filePatternsText)
        {
            if (string.IsNullOrWhiteSpace(filePatternsText))
            {
                return Array.Empty<string>();
            }

            return filePatternsText
                .Split('/')
                .Select(pattern => pattern.Trim())
                .Where(pattern => !string.IsNullOrEmpty(pattern))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
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

        private sealed class FilterCriteria
        {
            public string FilePath { get; set; } = string.Empty;
            public string FileName { get; set; } = string.Empty;
            public string Extension { get; set; } = string.Empty;
            public string RowNumber { get; set; } = string.Empty;
            public string GrepResult { get; set; } = string.Empty;
            public string Method { get; set; } = string.Empty;
        }

        private sealed class ResultColumnInfo
        {
            public ResultColumnInfo(string name, string headerText)
            {
                Name = name ?? string.Empty;
                HeaderText = headerText ?? string.Empty;
            }

            public string Name { get; }
            public string HeaderText { get; }
        }

        private sealed class SakuraGrepImportData
        {
            public List<SearchResult> Results { get; set; } = new List<SearchResult>();
            public SakuraSearchConditions SearchConditions { get; set; } = new SakuraSearchConditions();
            public bool HasSearchConditions { get; set; }
        }

        private sealed class SakuraSearchConditions
        {
            public DateTime ExportedAt { get; set; } = DateTime.Now;
            public string SearchFolder { get; set; } = string.Empty;
            public string FilePattern { get; set; } = string.Empty;
            public string Keyword { get; set; } = string.Empty;
            public string MultiKeywords { get; set; } = string.Empty;
            public string ExcludeFolders { get; set; } = string.Empty;
            public string ExcludeExtensions { get; set; } = string.Empty;
            public bool SearchSubDir { get; set; }
            public bool CaseSensitive { get; set; }
            public bool UseRegex { get; set; }
            public bool DeriveMethod { get; set; }
            public bool IgnoreComment { get; set; }
            public bool IgnoreBinaryFile { get; set; }
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
