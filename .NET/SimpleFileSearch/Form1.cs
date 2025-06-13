using System;
using System.Collections.Generic;
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

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // タイトルにバージョン情報を表示
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version}";

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
                        SearchFilesAndFoldersWithRegex(rootPath, regex, includeFolderNames, foundFiles);
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
                            Regex regex = new Regex(Regex.Escape(searchPattern), RegexOptions.IgnoreCase);
                            string rootPath = cmbFolderPath.Text;
                            SearchFilesAndFoldersWithRegex(rootPath, regex, includeFolderNames, foundFiles);
                        }
                        else
                        {
                            // 通常のワイルドカードモード
                            foundFiles.AddRange(Directory.GetFiles(cmbFolderPath.Text, searchPattern, SearchOption.AllDirectories));
            
                            // フォルダ名も検索対象に含める場合
                            if (includeFolderNames)
                            {
                                string rootPath = cmbFolderPath.Text;
                                SearchFoldersWithWildcard(rootPath, searchPattern, foundFiles);
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

                // 結果を表示
                foreach (string file in foundFiles)
                {
                    dataGridViewResults.Rows.Add(file);
                }

                // 結果件数を表示
                this.Text = $"シンプルなファイル検索 - {foundFiles.Count} 件のファイルが見つかりました";

                labelResult.Text = $"Result: {foundFiles.Count} hit";
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
        private void SearchFilesAndFoldersWithRegex(string folderPath, Regex regex, bool includeFolderNames, List<string> results)
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
                    SearchFilesAndFoldersWithRegex(dir, regex, includeFolderNames, results);
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
        private void SearchFoldersWithWildcard(string folderPath, string searchPattern, List<string> results)
        {
            try
            {
                // 指定されたパターンに合致するかテスト用の関数
                Func<string, bool> matchesPattern = (name) => 
                {
                    // ワイルドカードをRegexのパターンに変換
                    string regexPattern = "^" + Regex.Escape(searchPattern)
                        .Replace("\\*", ".*")
                        .Replace("\\?", ".") + "$";
                    
                    return Regex.IsMatch(name, regexPattern, RegexOptions.IgnoreCase);
                };

                // サブフォルダを処理
                foreach (string dir in Directory.GetDirectories(folderPath))
                {
                    string dirName = Path.GetFileName(dir);
                    
                    // フォルダ名がパターンに合致する場合
                    if (matchesPattern(dirName))
                    {
                        // フォルダが見つかった場合、そのフォルダ内のすべてのファイルを追加
                        results.AddRange(Directory.GetFiles(dir, "*.*", SearchOption.AllDirectories));
                    }
                    
                    // 再帰的に検索
                    SearchFoldersWithWildcard(dir, searchPattern, results);
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

        private void dataGridViewResults_CellDoubleClick ( object sender, DataGridViewCellEventArgs e )
        {
            if (e.RowIndex >= 0 && dataGridViewResults.Rows[e.RowIndex].Cells[0].Value != null)
            {
                string filePath = dataGridViewResults.Rows[e.RowIndex].Cells[0].Value.ToString();
                
                if (File.Exists(filePath))
                {
                    try
                    {
                        // エクスプローラーでフォルダを開いてファイルを選択
                        Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    UsePartialMatch = chkPartialMatch.Checked
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

                    // 最新の項目をテキストボックスに表示
                    if (cmbKeyword.Items.Count > 0)
                    {
                        cmbKeyword.SelectedIndex = 0;
                    }

                    if (cmbFolderPath.Items.Count > 0)
                    {
                        cmbFolderPath.SelectedIndex = 0;
                    }
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
    }
}