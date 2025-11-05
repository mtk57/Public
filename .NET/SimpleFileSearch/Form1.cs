using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class MainForm: Form
    {
        private const int MaxHistoryItems = 20;
        private const string SettingsFileName = "SimpleFileSearch.json";
        private List<SearchResult> _allResults = new List<SearchResult>();
        private List<SearchResult> _filteredResults = new List<SearchResult>();
        private string _currentSortColumn;
        private bool _sortAscending = true;
        private bool _suppressFilterEvent;

        public MainForm()
        {
            InitializeComponent();

            // ドラッグアンドドロップを有効にする
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);

            // フィルタ入力時に結果を反映
            txtFilePathFilter.TextChanged += FilterTextBox_TextChanged;
            txtFileNameFilter.TextChanged += FilterTextBox_TextChanged;
            txtExtFilter.TextChanged += FilterTextBox_TextChanged;

            // 列ヘッダークリックでソート
            dataGridViewResults.ColumnHeaderMouseClick += dataGridViewResults_ColumnHeaderMouseClick;
            dataGridViewResults.AutoGenerateColumns = false;
            foreach (DataGridViewColumn column in dataGridViewResults.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Programmatic;
            }

            btnDeleteByExt.Enabled = false;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // タイトルにバージョン情報を表示
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            // 設定ファイルからデータを読み込む
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 設定ファイルにデータを保存
            SaveSettings();
        }

        private void btnBrowse_Click ( object sender, EventArgs e )
        {
            // フォルダ選択ダイアログを表示
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "検索するフォルダを選択してください";
                
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    cmbFolderPath.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void btnSearch_Click ( object sender, EventArgs e )
        {
            // バリデーションチェック
            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbKeyword.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                MessageBox.Show("検索するフォルダを選択してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbFolderPath.Focus();
                return;
            }

            if (!Directory.Exists(cmbFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cmbFolderPath.Focus();
                return;
            }

            // 履歴を更新（キーワード）
            UpdateComboBoxHistory(cmbKeyword, cmbKeyword.Text);

            // 履歴を更新（フォルダパス）
            UpdateComboBoxHistory(cmbFolderPath, cmbFolderPath.Text);

            // 以前の検索結果をクリア
            dataGridViewResults.Rows.Clear();

            try
            {
                Cursor = Cursors.WaitCursor;

                string searchPattern = cmbKeyword.Text;
                bool useRegex = chkUseRegex.Checked;
                bool includeFolderNames = chkIncludeFolderNames.Checked;
                bool searchSubDir = chkSearchSubDir.Checked;
                
                List<string> foundFiles = new List<string>();

                // ファイル検索
                if (useRegex)
                {
                    // 正規表現モード
                    try
                    {
                        Regex regex = new Regex(searchPattern, RegexOptions.IgnoreCase);
                        
                        // ディレクトリを再帰的に列挙
                        string rootPath = cmbFolderPath.Text;
                        SearchFilesAndFoldersWithRegex(rootPath, regex, includeFolderNames, foundFiles, searchSubDir);
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show($"無効な正規表現です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    // 部分一致モードまたはワイルドカードモード
                    bool usePartialMatch = chkPartialMatch.Checked;

                    try
                    {
                        if (usePartialMatch)
                        {
                            // 部分一致モード - 正規表現に変換して検索
                            Regex regex = CreateWildcardRegex(searchPattern, anchorMatch: false);
                            string rootPath = cmbFolderPath.Text;
                            SearchFilesAndFoldersWithRegex(rootPath, regex, includeFolderNames, foundFiles, searchSubDir);
                        }
                        else
                        {
                            // 通常のワイルドカードモード
                            SearchOption searchOption = searchSubDir ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                            foundFiles.AddRange(Directory.GetFiles(cmbFolderPath.Text, searchPattern, searchOption));

                            // フォルダ名も検索対象に含める場合
                            if (includeFolderNames)
                            {
                                string rootPath = cmbFolderPath.Text;
                                SearchFoldersWithWildcard(rootPath, searchPattern, foundFiles, searchSubDir);
                            }
                        }
                    }
                    catch (ArgumentException ex)
                    {
                        MessageBox.Show($"無効な検索パターンです: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 重複を削除してソート
                foundFiles = foundFiles.Distinct().OrderBy(f => f).ToList();

                // 検索結果リストを構築
                List<SearchResult> results = new List<SearchResult>();
                foreach (string file in foundFiles)
                {
                    try
                    {
                        if (!File.Exists(file))
                        {
                            continue;
                        }

                        FileInfo info = new FileInfo(file);
                        results.Add(new SearchResult
                        {
                            FilePath = file,
                            FileName = Path.GetFileName(file),
                            Extension = info.Extension?.TrimStart('.'),
                            Size = info.Length,
                            LastWriteTime = info.LastWriteTime
                        });
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // アクセスできないファイルは無視
                    }
                    catch (Exception)
                    {
                        // その他の例外も無視して次へ
                    }
                }

                _allResults = results;
                _currentSortColumn = columnFilePath.Name;
                _sortAscending = true;
                ClearFilters();
                ApplyFilterAndSort();

                // 結果件数を表示
                //this.Text = $"シンプルなファイル検索 - {foundFiles.Count} 件のファイルが見つかりました";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        // 正規表現を使用してファイルとフォルダを再帰的に検索
        private void SearchFilesAndFoldersWithRegex(string folderPath, Regex regex, bool includeFolderNames, List<string> results, bool searchSubDir)
        {
            try
            {
                // ファイルを検索
                foreach (string file in Directory.GetFiles(folderPath))
                {
                    string fileName = Path.GetFileName(file);
                    if (regex.IsMatch(fileName))
                    {
                        results.Add(file);
                    }
                }

                // サブフォルダを処理
                if (searchSubDir)
                {
                    foreach (string dir in Directory.GetDirectories(folderPath))
                    {
                        // フォルダ名を検索対象に含める場合
                        if (includeFolderNames)
                        {
                            string dirName = Path.GetFileName(dir);
                            if (regex.IsMatch(dirName))
                            {
                                // フォルダが見つかった場合、そのフォルダ内のすべてのファイルを追加
                                results.AddRange(Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories));
                            }
                        }

                        // 再帰的に検索
                        SearchFilesAndFoldersWithRegex(dir, regex, includeFolderNames, results, searchSubDir);
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // アクセス権限がない場合は無視
            }
            catch (Exception)
            {
                // その他のエラーも無視して次へ
            }
        }

        // ワイルドカードを使用してフォルダを検索
        private void SearchFoldersWithWildcard(string folderPath, string searchPattern, List<string> results, bool searchSubDir)
        {
            try
            {
                Regex folderRegex = CreateWildcardRegex(searchPattern, anchorMatch: true);

                // サブフォルダを処理
                if (searchSubDir)
                {
                    foreach (string dir in Directory.GetDirectories(folderPath))
                    {
                        string dirName = Path.GetFileName(dir);

                        // フォルダ名がパターンに合致する場合
                        if (folderRegex.IsMatch(dirName))
                        {
                            // フォルダが見つかった場合、そのフォルダ内のすべてのファイルを追加
                            results.AddRange(Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories));
                        }

                        // 再帰的に検索
                        SearchFoldersWithWildcard(dir, searchPattern, results, searchSubDir);
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // アクセス権限がない場合は無視
            }
            catch (Exception)
            {
                // その他のエラーも無視して次へ
            }
        }

        private static Regex CreateWildcardRegex(string pattern, bool anchorMatch)
        {
            string regexPattern = ConvertWildcardToRegexPattern(pattern, anchorMatch);
            return new Regex(regexPattern, RegexOptions.IgnoreCase);
        }

        private static string ConvertWildcardToRegexPattern(string pattern, bool anchorMatch)
        {
            string escaped = Regex.Escape(pattern ?? string.Empty)
                .Replace("\\*", ".*")
                .Replace("\\?", ".");

            if (anchorMatch)
            {
                return "^" + escaped + "$";
            }

            return escaped;
        }

        private void FilterTextBox_TextChanged(object sender, EventArgs e)
        {
            if (_suppressFilterEvent)
            {
                return;
            }

            ApplyFilterAndSort();
        }

        private void dataGridViewResults_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var column = dataGridViewResults.Columns[e.ColumnIndex];
            if (column == null || string.IsNullOrEmpty(column.Name))
            {
                return;
            }

            if (_currentSortColumn == column.Name)
            {
                _sortAscending = !_sortAscending;
            }
            else
            {
                _currentSortColumn = column.Name;
                _sortAscending = true;
            }

            ApplyFilterAndSort();
        }

        private void ApplyFilterAndSort()
        {
            IEnumerable<SearchResult> query = _allResults ?? Enumerable.Empty<SearchResult>();

            string filePathFilter = txtFilePathFilter.Text.Trim();
            if (!string.IsNullOrEmpty(filePathFilter))
            {
                query = query.Where(r => r.FilePath.IndexOf(filePathFilter, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            string fileNameFilter = txtFileNameFilter.Text.Trim();
            if (!string.IsNullOrEmpty(fileNameFilter))
            {
                query = query.Where(r => r.FileName.IndexOf(fileNameFilter, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            string extFilter = txtExtFilter.Text.Trim();
            if (!string.IsNullOrEmpty(extFilter))
            {
                string normalized = extFilter.TrimStart('.');
                query = query.Where(r => (r.Extension ?? string.Empty).IndexOf(normalized, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            if (string.IsNullOrEmpty(_currentSortColumn))
            {
                _currentSortColumn = columnFilePath.Name;
            }

            IOrderedEnumerable<SearchResult> ordered;
            switch (_currentSortColumn)
            {
                case nameof(clmFileName):
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.FileName, StringComparer.OrdinalIgnoreCase)
                        : query.OrderByDescending(r => r.FileName, StringComparer.OrdinalIgnoreCase);
                    break;
                case nameof(clmExt):
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.Extension ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                        : query.OrderByDescending(r => r.Extension ?? string.Empty, StringComparer.OrdinalIgnoreCase);
                    break;
                case nameof(clmSize):
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.Size)
                        : query.OrderByDescending(r => r.Size);
                    break;
                case nameof(clmUpdateDate):
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.LastWriteTime)
                        : query.OrderByDescending(r => r.LastWriteTime);
                    break;
                case nameof(columnFilePath):
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.FilePath, StringComparer.OrdinalIgnoreCase)
                        : query.OrderByDescending(r => r.FilePath, StringComparer.OrdinalIgnoreCase);
                    break;
                default:
                    ordered = _sortAscending
                        ? query.OrderBy(r => r.FilePath, StringComparer.OrdinalIgnoreCase)
                        : query.OrderByDescending(r => r.FilePath, StringComparer.OrdinalIgnoreCase);
                    break;
            }

            List<SearchResult> displayList = ordered.ToList();
            _filteredResults = displayList;

            dataGridViewResults.SuspendLayout();
            dataGridViewResults.Rows.Clear();
            foreach (SearchResult result in displayList)
            {
                dataGridViewResults.Rows.Add(
                    result.FilePath,
                    result.FileName,
                    result.Extension,
                    result.Size.ToString("N0"),
                    result.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss"));
            }

            foreach (DataGridViewColumn column in dataGridViewResults.Columns)
            {
                column.HeaderCell.SortGlyphDirection = SortOrder.None;
            }

            if (!string.IsNullOrEmpty(_currentSortColumn))
            {
                DataGridViewColumn targetColumn = dataGridViewResults.Columns[_currentSortColumn];
                if (targetColumn != null)
                {
                    targetColumn.HeaderCell.SortGlyphDirection = _sortAscending ? SortOrder.Ascending : SortOrder.Descending;
                }
            }

            dataGridViewResults.ResumeLayout();

            int filteredCount = displayList.Count;
            int totalCount = _allResults?.Count ?? 0;
            btnDeleteByExt.Enabled = filteredCount > 0;

            this.Text = $"シンプルなファイル検索 - {totalCount} 件のファイルが見つかりました";
            if (filteredCount == totalCount)
            {
                labelResult.Text = $"Result: {filteredCount} hit";
            }
            else
            {
                labelResult.Text = $"Result: {filteredCount} hit (total {totalCount})";
            }
        }

        private void ClearFilters()
        {
            _suppressFilterEvent = true;
            try
            {
                txtFilePathFilter.Clear();
                txtFileNameFilter.Clear();
                txtExtFilter.Clear();
            }
            finally
            {
                _suppressFilterEvent = false;
            }
        }

        private static void RestoreComboBoxText(ComboBox comboBox, string storedValue)
        {
            if (comboBox == null)
            {
                return;
            }

            if (storedValue == null)
            {
                if (comboBox.Items.Count > 0)
                {
                    comboBox.SelectedIndex = 0;
                }
                else
                {
                    comboBox.SelectedIndex = -1;
                    comboBox.Text = string.Empty;
                }
                return;
            }

            if (storedValue.Length == 0)
            {
                comboBox.SelectedIndex = -1;
                comboBox.Text = string.Empty;
                return;
            }

            int index = comboBox.Items.IndexOf(storedValue);
            if (index >= 0)
            {
                comboBox.SelectedIndex = index;
            }
            else
            {
                comboBox.SelectedIndex = -1;
                comboBox.Text = storedValue;
            }
        }

        private void dataGridViewResults_CellDoubleClick ( object sender, DataGridViewCellEventArgs e )
        {
            if (e.RowIndex >= 0 && dataGridViewResults.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string filePath = dataGridViewResults.Rows[e.RowIndex].Cells[0].Value.ToString();
                
                if (File.Exists(filePath))
                {
                    try
                    {
                        if (chkDblClickToOpen.Checked)
                        {
                            // ファイルを開く
                            Process.Start(filePath);
                        }
                        else
                        {
                            // エクスプローラーでフォルダを開いてファイルを選択
                            Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void cmbKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enterキーが押された場合、検索ボタンクリックイベントを呼び出す
                btnSearch_Click(this, EventArgs.Empty);
                // キーイベントを処理済みとしてマーク
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        #region コンボボックス履歴の管理

        private void UpdateComboBoxHistory(ComboBox comboBox, string item)
        {
            // 既に同じ項目が存在する場合は削除（重複を避けるため）
            if (comboBox.Items.Contains(item))
            {
                comboBox.Items.Remove(item);
            }

            // 最大数に達している場合は最も古い項目を削除
            if (comboBox.Items.Count >= MaxHistoryItems)
            {
                comboBox.Items.RemoveAt(comboBox.Items.Count - 1);
            }

            // 新しい項目を先頭に追加
            comboBox.Items.Insert(0, item);
            comboBox.Text = item;
        }

        #endregion

        #region 設定の保存と読み込み

        private void SaveSettings()
        {
            try
            {
                AppSettings settings = new AppSettings
                {
                    KeywordHistory = new List<string>(),
                    FolderPathHistory = new List<string>(),
                    UseRegex = chkUseRegex.Checked,
                    IncludeFolderNames = chkIncludeFolderNames.Checked,
                    UsePartialMatch = chkPartialMatch.Checked,
                    SearchSubDir = chkSearchSubDir.Checked,
                    DblClickToOpen = chkDblClickToOpen.Checked
                };

                // キーワード履歴を保存
                foreach (var item in cmbKeyword.Items)
                {
                    settings.KeywordHistory.Add(item.ToString());
                }

                // フォルダパス履歴を保存
                foreach (var item in cmbFolderPath.Items)
                {
                    settings.FolderPathHistory.Add(item.ToString());
                }

                settings.LastKeyword = cmbKeyword.Text ?? string.Empty;
                settings.LastFolderPath = cmbFolderPath.Text ?? string.Empty;

                // JavaScriptSerializerを使用してJSONに変換
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(settings);
                
                // JSONファイルとして保存
                File.WriteAllText(GetSettingsFilePath(), json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // 保存エラーは無視（ユーザーに通知しない）
                Console.WriteLine($"設定の保存中にエラーが発生しました: {ex.Message}");
            }
        }

        private void LoadSettings()
        {
            string settingsFilePath = GetSettingsFilePath();
            
            if (!File.Exists(settingsFilePath))
            {
                return; // 設定ファイルが存在しない場合は何もしない
            }

            try
            {
                string json = File.ReadAllText(settingsFilePath, Encoding.UTF8);
                
                // JavaScriptSerializerを使用してJSONをデシリアライズ
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                AppSettings settings = serializer.Deserialize<AppSettings>(json);

                if (settings != null)
                {
                    // キーワード履歴を読み込み
                    cmbKeyword.Items.Clear();
                    foreach (var item in settings.KeywordHistory)
                    {
                        cmbKeyword.Items.Add(item);
                    }

                    // フォルダパス履歴を読み込み
                    cmbFolderPath.Items.Clear();
                    foreach (var item in settings.FolderPathHistory)
                    {
                        cmbFolderPath.Items.Add(item);
                    }

                    // 正規表現の設定を読み込み
                    chkUseRegex.Checked = settings.UseRegex;

                    // フォルダ名検索の設定を読み込み
                    chkIncludeFolderNames.Checked = settings.IncludeFolderNames;

                    // 部分一致モードの設定を読み込み (追加)
                    chkPartialMatch.Checked = settings.UsePartialMatch;

                    // サブフォルダ検索の設定を読み込み
                    chkSearchSubDir.Checked = settings.SearchSubDir;

                    // ダブルクリックでファイルを開くの設定を読み込み
                    chkDblClickToOpen.Checked = settings.DblClickToOpen;

                    // 最新の入力値を再現
                    RestoreComboBoxText(cmbKeyword, settings.LastKeyword);
                    RestoreComboBoxText(cmbFolderPath, settings.LastFolderPath);
                }
            }
            catch (Exception ex)
            {
                // 読み込みエラーは無視（ユーザーに通知しない）
                Console.WriteLine($"設定の読み込み中にエラーが発生しました: {ex.Message}");
            }
        }

        private string GetSettingsFilePath()
        {
            return Path.Combine(Application.StartupPath, SettingsFileName);
        }

        #endregion

        #region ドラッグアンドドロップ

        private void cmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグされたものがファイルまたはフォルダの場合、カーソルを変更
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
            // ドロップされたファイルまたはフォルダのパスを取得
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (files.Length > 0)
            {
                // 最初のアイテムのパスを使用
                string path = files[0];

                // それがフォルダであるか、ファイルであるかを確認
                if (Directory.Exists(path))
                {
                    // フォルダの場合はそのパスを直接使用
                    cmbFolderPath.Text = path;
                }
                else if (File.Exists(path))
                {
                    // ファイルの場合は、そのファイルが含まれるフォルダのパスを使用
                    cmbFolderPath.Text = Path.GetDirectoryName(path);
                }
            }
        }

        #endregion

        private void btnFileCopy_Click ( object sender, EventArgs e )
        {
            try
            {
                // 選択済みの行を取得（行選択が無い場合はセル選択から補完）
                var selectedRows = dataGridViewResults.SelectedRows
                    .Cast<DataGridViewRow>()
                    .ToList();

                if (selectedRows.Count == 0)
                {
                    selectedRows = dataGridViewResults.SelectedCells
                        .Cast<DataGridViewCell>()
                        .Select(cell => dataGridViewResults.Rows[cell.RowIndex])
                        .Distinct()
                        .ToList();
                }

                if (selectedRows.Count == 0)
                {
                    MessageBox.Show("コピーするファイルを選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                StringCollection fileList = new StringCollection();

                foreach (var row in selectedRows)
                {
                    string filePath = row.Cells[0].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                    {
                        fileList.Add(filePath);
                    }
                }

                if (fileList.Count == 0)
                {
                    MessageBox.Show("コピー可能なファイルが選択されていません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Clipboard.SetFileDropList(fileList);
            }
            catch (System.Runtime.InteropServices.ExternalException ex)
            {
                MessageBox.Show($"クリップボードにコピーできませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"コピー処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteByExt_Click ( object sender, EventArgs e )
        {
            if (_filteredResults == null || _filteredResults.Count == 0)
            {
                return;
            }

            using (var dialog = new DeleteByExtensionForm(_filteredResults))
            {
                dialog.ShowDialog(this);
                IReadOnlyList<string> deletedFiles = dialog.DeletedFiles;
                if (deletedFiles != null && deletedFiles.Count > 0)
                {
                    RemoveDeletedFiles(deletedFiles);
                }
            }
        }

        private void RemoveDeletedFiles(IReadOnlyCollection<string> deletedFiles)
        {
            if (deletedFiles == null || deletedFiles.Count == 0 || _allResults == null)
            {
                return;
            }

            HashSet<string> deletedSet = new HashSet<string>(deletedFiles, StringComparer.OrdinalIgnoreCase);
            _allResults = _allResults
                .Where(r => !deletedSet.Contains(r.FilePath))
                .ToList();

            // 表示から削除されたファイルを除外するため再描画
            ApplyFilterAndSort();
        }
    }
}
