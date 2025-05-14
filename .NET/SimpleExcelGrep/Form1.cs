using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
// Drawing 名前空間の型を明確にするために using エイリアスディレクティブ を使用することが推奨されます
// using DDrawing = DocumentFormat.OpenXml.Drawing;
// using DSSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SimpleExcelGrep
{
    public partial class MainForm : Form
    {
        private CancellationTokenSource _cancellationTokenSource;
        private string _settingsFilePath = "settings.json";
        private bool _isSearching = false;
        private const int MaxHistoryItems = 10;
        private List<SearchResult> _searchResults = new List<SearchResult>(); // 検索結果を保持するリスト
        private bool _isLoading = false;

        // ログファイルのパス
        private string _logFilePath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            "SimpleExcelGrep_log.txt");

        // モデルクラス - 設定の保存/読み込み用
        [DataContract]
        public class Settings
        {
            [DataMember]
            public string FolderPath { get; set; } = "";

            [DataMember]
            public string SearchKeyword { get; set; } = "";

            [DataMember]
            public bool UseRegex { get; set; } = false;

            [DataMember]
            public string IgnoreKeywords { get; set; } = "";

            [DataMember]
            public List<string> FolderPathHistory { get; set; } = new List<string>();

            [DataMember]
            public List<string> SearchKeywordHistory { get; set; } = new List<string>();

            [DataMember]
            public List<string> IgnoreKeywordsHistory { get; set; } = new List<string>();

            [DataMember]
            public bool RealTimeDisplay { get; set; } = true; // リアルタイム表示設定

            [DataMember]
            public int MaxParallelism { get; set; } = Environment.ProcessorCount; // 並列処理数

            [DataMember]
            public bool FirstHitOnly { get; set; } = false; // 追加: 最初のヒットのみ検索設定

            [DataMember]
            public bool SearchShapes { get; set; } = false; // 図形内の文字列を検索するかどうか
        }

        // 検索結果を格納するクラス
        public class SearchResult
        {
            public string FilePath { get; set; }
            public string SheetName { get; set; }
            public string CellPosition { get; set; } // セル位置または "図形内" など
            public string CellValue { get; set; }    // セルの値または図形内のテキスト
        }

        public MainForm()
        {
            InitializeComponent();
            this.FormClosing += MainForm_FormClosing;
            this.Load += MainForm_Load;
            btnSelectFolder.Click += BtnSelectFolder_Click;
            btnStartSearch.Click += BtnStartSearch_Click;
            btnCancelSearch.Click += BtnCancelSearch_Click;
            grdResults.DoubleClick += GrdResults_DoubleClick;

            // グリッドにキー押下イベントを追加
            grdResults.KeyDown += GrdResults_KeyDown;

            // 複数行選択を可能にする
            grdResults.MultiSelect = true;
            grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // コンテキストメニューを追加
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("コピー");
            copyMenuItem.Click += (s, e) => CopySelectedRowsToClipboard();
            contextMenu.Items.Add(copyMenuItem);

            ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("すべて選択");
            selectAllMenuItem.Click += (s, e) => grdResults.SelectAll();
            contextMenu.Items.Add(selectAllMenuItem);

            grdResults.ContextMenuStrip = contextMenu;

            // リアルタイム表示チェックボックスの状態変更イベントを登録
            chkRealTimeDisplay.CheckedChanged += (s, e) => {
                if (!_isLoading)
                {
                    // 設定を保存
                    SaveSettings();
                }
            };

            // 並列数設定の変更イベントを登録
            nudParallelism.ValueChanged += (s, e) => {
                // 読み込み中は保存をスキップ
                if (!_isLoading)
                {
                    SaveSettings();
                }
            };

            // 最初のヒットのみチェックボックスの状態変更イベントを登録
            chkFirstHitOnly.CheckedChanged += (s, e) => {
                if (!_isLoading)
                {
                    // 設定を保存
                    SaveSettings();
                }
            };

            // 図形内検索チェックボックスの状態変更イベントを登録
            chkSearchShapes.CheckedChanged += (s, e) => {
                if (!_isLoading)
                {
                    // 設定を保存
                    SaveSettings();
                }
            };
        }

        // ログを記録するメソッド
        private void LogMessage(string message, bool showInStatus = false)
        {
            //string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //string logEntry = $"[{timestamp}] {message}";

            //try
            //{
            //    // ログファイルに追記
            //    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);

            //    // デバッグコンソールにも出力
            //    Debug.WriteLine(logEntry);

            //    // オプションでステータスバーに表示
            //    if (showInStatus)
            //    {
            //        UpdateStatus(message);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    // ログ出力自体が失敗した場合、ステータスバーだけに表示
            //    UpdateStatus($"ログの記録に失敗: {ex.Message}");
            //}
        }

        // 環境情報取得メソッド（Late Binding方式）
        private void LogEnvironmentInfo()
        {
            LogMessage("======== 環境情報 ========");
            LogMessage($"OSバージョン: {Environment.OSVersion}");
            LogMessage($".NETバージョン: {Environment.Version}");
            LogMessage($"64ビットOS: {Environment.Is64BitOperatingSystem}");
            LogMessage($"64ビットプロセス: {Environment.Is64BitProcess}");
            LogMessage($"マシン名: {Environment.MachineName}");

            // Officeのバージョン情報取得を試行（Late Binding方式）
            try
            {
                // Excelアプリケーションを作成
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType != null)
                {
                    object excelApp = Activator.CreateInstance(excelType);

                    // バージョン情報を取得
                    object version = excelType.InvokeMember("Version",
                        System.Reflection.BindingFlags.GetProperty,
                        null, excelApp, null);
                    LogMessage($"Excel バージョン: {version}");

                    // ビルド情報を取得
                    object build = excelType.InvokeMember("Build",
                        System.Reflection.BindingFlags.GetProperty,
                        null, excelApp, null);
                    LogMessage($"Excel ビルド: {build}");

                    // COMの早期解放
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); } catch { }
                }
                else
                {
                    LogMessage("Excel.ApplicationのCOMタイプが見つかりません");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Excelバージョン取得エラー: {ex.Message}");
            }

            LogMessage("==========================");
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LogMessage("アプリケーション起動");
            LogEnvironmentInfo();
            LoadSettings();
        }

        private void LoadSettings()
        {
            _isLoading = true;

            LogMessage("設定の読み込みを開始");
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    // DataContractJsonSerializerを使用してJSON読み込み
                    using (FileStream fs = new FileStream(_settingsFilePath, FileMode.Open))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                        Settings settings = (Settings)serializer.ReadObject(fs);

                        // コンボボックスの履歴をクリアして設定
                        cmbFolderPath.Items.Clear();
                        foreach (var path in settings.FolderPathHistory)
                        {
                            cmbFolderPath.Items.Add(path);
                        }
                        cmbFolderPath.Text = settings.FolderPath;

                        // 検索キーワード履歴
                        cmbKeyword.Items.Clear();
                        foreach (var keyword in settings.SearchKeywordHistory)
                        {
                            cmbKeyword.Items.Add(keyword);
                        }
                        cmbKeyword.Text = settings.SearchKeyword;

                        // 無視キーワード履歴
                        cmbIgnoreKeywords.Items.Clear();
                        foreach (var keyword in settings.IgnoreKeywordsHistory)
                        {
                            cmbIgnoreKeywords.Items.Add(keyword);
                        }
                        cmbIgnoreKeywords.Text = settings.IgnoreKeywords;

                        // その他の設定
                        chkRegex.Checked = settings.UseRegex;

                        // リアルタイム表示設定
                        chkRealTimeDisplay.Checked = settings.RealTimeDisplay;

                        // 並列処理設定
                        nudParallelism.Value = Math.Min(Math.Max(settings.MaxParallelism, 1), 32);

                        // 最初のヒットのみ設定
                        chkFirstHitOnly.Checked = settings.FirstHitOnly;

                        // 図形内検索設定
                        chkSearchShapes.Checked = settings.SearchShapes;

                        LogMessage("設定の読み込みが成功しました");
                    }
                }
                else
                {
                    LogMessage("設定ファイルが見つかりません");
                }
            }
            catch (Exception ex)
            {
                // 読み込み失敗時は何もしない
                LogMessage($"設定の読み込みに失敗: {ex.Message}");
            }
            finally
            {
                _isLoading = false;
            }
        }

        private void SaveSettings()
        {
            LogMessage("設定の保存を開始");
            try
            {
                // フォルダパス履歴を更新
                List<string> folderPathHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbFolderPath.Text))
                {
                    folderPathHistory.Add(cmbFolderPath.Text);
                }
                foreach (var item in cmbFolderPath.Items)
                {
                    string path = item.ToString();
                    if (!folderPathHistory.Contains(path) && folderPathHistory.Count < MaxHistoryItems)
                    {
                        folderPathHistory.Add(path);
                    }
                }

                // 検索キーワード履歴を更新
                List<string> searchKeywordHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbKeyword.Text))
                {
                    searchKeywordHistory.Add(cmbKeyword.Text);
                }
                foreach (var item in cmbKeyword.Items)
                {
                    string keyword = item.ToString();
                    if (!searchKeywordHistory.Contains(keyword) && searchKeywordHistory.Count < MaxHistoryItems)
                    {
                        searchKeywordHistory.Add(keyword);
                    }
                }

                // 無視キーワード履歴を更新
                List<string> ignoreKeywordsHistory = new List<string>();
                if (!string.IsNullOrEmpty(cmbIgnoreKeywords.Text))
                {
                    ignoreKeywordsHistory.Add(cmbIgnoreKeywords.Text);
                }
                foreach (var item in cmbIgnoreKeywords.Items)
                {
                    string keyword = item.ToString();
                    if (!ignoreKeywordsHistory.Contains(keyword) && ignoreKeywordsHistory.Count < MaxHistoryItems)
                    {
                        ignoreKeywordsHistory.Add(keyword);
                    }
                }

                // 設定を保存
                Settings settings = new Settings
                {
                    FolderPath = cmbFolderPath.Text,
                    SearchKeyword = cmbKeyword.Text,
                    UseRegex = chkRegex.Checked,
                    IgnoreKeywords = cmbIgnoreKeywords.Text,
                    FolderPathHistory = folderPathHistory,
                    SearchKeywordHistory = searchKeywordHistory,
                    IgnoreKeywordsHistory = ignoreKeywordsHistory,
                    RealTimeDisplay = chkRealTimeDisplay.Checked,
                    MaxParallelism = (int)nudParallelism.Value,
                    FirstHitOnly = chkFirstHitOnly.Checked,
                    SearchShapes = chkSearchShapes.Checked
                };

                // DataContractJsonSerializerを使用してJSON保存
                using (FileStream fs = new FileStream(_settingsFilePath, FileMode.Create))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                    serializer.WriteObject(fs, settings);
                }

                LogMessage("設定の保存が成功しました");
            }
            catch (Exception ex)
            {
                LogMessage($"設定の保存に失敗: {ex.Message}");
                MessageBox.Show($"設定の保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 履歴コンボボックスに項目を追加（重複なしで先頭に配置）
        private void AddToComboBoxHistory(ComboBox comboBox, string item)
        {
            if (string.IsNullOrEmpty(item))
                return;

            // 既存の項目を削除（重複を防ぐ）
            if (comboBox.Items.Contains(item))
            {
                comboBox.Items.Remove(item);
            }

            // 先頭に追加
            comboBox.Items.Insert(0, item);

            // 最大履歴数を超えた場合、古い項目を削除
            while (comboBox.Items.Count > MaxHistoryItems)
            {
                comboBox.Items.RemoveAt(comboBox.Items.Count - 1);
            }

            // 現在の選択項目を設定
            comboBox.Text = item;
        }

        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            LogMessage("フォルダ選択ボタンがクリックされました");
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "検索するフォルダを選択してください";
                dialog.ShowNewFolderButton = false;

                if (!string.IsNullOrEmpty(cmbFolderPath.Text) && Directory.Exists(cmbFolderPath.Text))
                {
                    dialog.SelectedPath = cmbFolderPath.Text;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    LogMessage($"選択されたフォルダ: {dialog.SelectedPath}");
                    AddToComboBoxHistory(cmbFolderPath, dialog.SelectedPath);
                }
                else
                {
                    LogMessage("フォルダ選択はキャンセルされました");
                }
            }
        }

        private async void BtnStartSearch_Click(object sender, EventArgs e)
        {
            LogMessage("検索開始ボタンがクリックされました");

            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                LogMessage("フォルダパスが入力されていません");
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(cmbFolderPath.Text))
            {
                LogMessage($"指定されたフォルダが存在しません: {cmbFolderPath.Text}");
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                LogMessage("検索キーワードが入力されていません");
                MessageBox.Show("検索キーワードを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 検索キーワードを履歴に追加
            LogMessage($"検索キーワード: {cmbKeyword.Text}");
            AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);

            // 無視キーワードを履歴に追加
            LogMessage($"無視キーワード: {cmbIgnoreKeywords.Text}");
            AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);

            // UIを検索中の状態に変更
            SetSearchingState(true);

            // キャンセルトークンを作成
            _cancellationTokenSource = new CancellationTokenSource();

            // 検索結果リストをクリア
            _searchResults.Clear();

            try
            {
                // 結果グリッドをクリア
                grdResults.Rows.Clear();

                // 無視キーワードのリストを作成
                List<string> ignoreKeywords = cmbIgnoreKeywords.Text
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(k => k.Trim())
                    .Where(k => !string.IsNullOrEmpty(k))
                    .ToList();

                LogMessage($"無視キーワードリスト: {string.Join(", ", ignoreKeywords)}");

                // 正規表現オブジェクト
                Regex regex = null;
                if (chkRegex.Checked)
                {
                    try
                    {
                        LogMessage($"正規表現を使用: {cmbKeyword.Text}");
                        regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled | RegexOptions.IgnoreCase); // 大文字小文字を区別しないオプションを追加
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"正規表現エラー: {ex.Message}");
                        MessageBox.Show($"正規表現が無効です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetSearchingState(false);
                        return;
                    }
                }

                // 検索処理を実行
                bool isRealTimeDisplay = chkRealTimeDisplay.Checked;
                LogMessage($"リアルタイム表示: {isRealTimeDisplay}, 最初のヒットのみ: {chkFirstHitOnly.Checked}, 図形内検索: {chkSearchShapes.Checked}, 並列数: {nudParallelism.Value}");

                // 検索結果を取得
                List<SearchResult> results = await SearchExcelFilesAsync(
                    cmbFolderPath.Text,
                    cmbKeyword.Text,
                    chkRegex.Checked,
                    regex,
                    ignoreKeywords,
                    isRealTimeDisplay,
                    chkSearchShapes.Checked, // 図形内検索フラグを渡す
                    _cancellationTokenSource.Token);

                // リアルタイム表示がOFFの場合または検索が途中でキャンセルされた場合に、
                // 最終的な結果をまとめて表示
                if (!isRealTimeDisplay || (_cancellationTokenSource != null && _cancellationTokenSource.IsCancellationRequested))
                {
                    LogMessage($"最終結果をまとめて表示: {_searchResults.Count}件");
                    DisplaySearchResults(_searchResults);
                }

                lblStatus.Text = $"検索完了: {_searchResults.Count} 件見つかりました";
                LogMessage($"検索完了: {_searchResults.Count} 件見つかりました");
            }
            catch (OperationCanceledException)
            {
                LogMessage($"検索はキャンセルされました: {_searchResults.Count} 件見つかりました");
                lblStatus.Text = $"検索は中止されました: {_searchResults.Count} 件見つかりました";

                // キャンセル時は現在までの結果を表示
                DisplaySearchResults(_searchResults);
            }
            catch (Exception ex)
            {
                LogMessage($"検索中にエラーが発生しました: {ex.Message}");
                MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラーが発生しました";

                // エラー時は現在までの結果を表示
                DisplaySearchResults(_searchResults);
            }
            finally
            {
                // UIを通常の状態に戻す
                SetSearchingState(false);
            }
        }

        private void cmbKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                // Enterキーが押された場合、検索ボタンクリックイベントを呼び出す
                BtnStartSearch_Click(this, EventArgs.Empty);
                // キーイベントを処理済みとしてマーク
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // 検索結果をグリッドに表示するメソッド
        private void DisplaySearchResults(List<SearchResult> results)
        {
            if (this.InvokeRequired)
            {
                LogMessage($"DisplaySearchResults: UIスレッドへの呼び出しが必要です。結果件数={results.Count}");
                this.Invoke(new Action(() => DisplaySearchResults(results)));
                return;
            }

            LogMessage($"DisplaySearchResults: グリッドをクリアして結果を表示します。件数={results.Count}");

            // 結果グリッドをクリア
            grdResults.Rows.Clear();

            // 結果をグリッドに表示
            foreach (var result in results)
            {
                string fileName = System.IO.Path.GetFileName(result.FilePath);
                grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);
                LogMessage($"グリッドに行を追加: {result.SheetName}, {result.CellPosition}, {result.CellValue}");
            }

            LogMessage($"グリッドに {results.Count} 件の結果を表示しました");
        }

        // 検索結果を1件追加するメソッド（リアルタイム表示用）
        private void AddSearchResult(SearchResult result)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => AddSearchResult(result)));
                return;
            }

            string fileName = System.IO.Path.GetFileName(result.FilePath);
            grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);

            // 最新の行にスクロール
            if (grdResults.Rows.Count > 0)
            {
                grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
            }
        }

        private void BtnCancelSearch_Click(object sender, EventArgs e)
        {
            LogMessage("検索キャンセルボタンがクリックされました");
            if (_isSearching && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
                lblStatus.Text = "キャンセル処理中...";
                LogMessage("検索をキャンセルしました");
            }
        }

        private void GrdResults_DoubleClick(object sender, EventArgs e)
        {
            LogMessage("GridView ダブルクリックイベント開始");

            if (grdResults.SelectedRows.Count <= 0)
            {
                LogMessage("選択行がありません");
                return;
            }

            string filePath = grdResults.SelectedRows[0].Cells["colFilePath"].Value?.ToString();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                LogMessage($"ファイルが見つかりません: {filePath}");
                MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Shiftキーが押されているかチェック
            bool isShiftPressed = (System.Windows.Forms.Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            LogMessage($"Shiftキー押下状態: {isShiftPressed}");

            if (isShiftPressed)
            {
                // Shift+ダブルクリック：フォルダを開く
                try
                {
                    string folderPath = System.IO.Path.GetDirectoryName(filePath);
                    LogMessage($"フォルダを開きます: {folderPath}");
                    System.Diagnostics.Process.Start("explorer.exe", folderPath);
                    lblStatus.Text = $"{folderPath} フォルダを開きました";
                }
                catch (Exception ex)
                {
                    LogMessage($"フォルダを開く際のエラー: {ex.Message}");
                    MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // 通常のダブルクリック：Excelファイルを開く
                string sheetName = grdResults.SelectedRows[0].Cells["colSheetName"].Value?.ToString();
                string cellPosition = grdResults.SelectedRows[0].Cells["colCellPosition"].Value?.ToString();
                LogMessage($"Excel操作を開始します: ファイル={filePath}, シート={sheetName}, セル={cellPosition}");

                // 図形内の場合はセル位置指定なしで開く
                if (cellPosition == "図形内" || cellPosition == "図形内 (GF)")
                {
                    OpenExcelFile(filePath, sheetName, null);
                }
                else
                {
                    OpenExcelFile(filePath, sheetName, cellPosition);
                }
            }
        }

        // Excelファイルを開くメソッド
        private void OpenExcelFile(string filePath, string sheetName, string cellPosition)
        {
            LogMessage($"OpenExcelFile 開始: ファイル={filePath}, シート={sheetName}, セル={cellPosition}");

            try
            {
                // Excel Interop を使用して開く
                LogMessage("Excel Interop を使用して開こうとしています...");
                bool interopSuccess = OpenExcelWithInterop(filePath, sheetName, cellPosition);
                LogMessage($"Excel Interop 結果: {(interopSuccess ? "成功" : "失敗")}");

                if (interopSuccess)
                {
                    return; // 成功したら終了
                }

                // 最も基本的な方法：ファイルを直接開く
                LogMessage("通常の方法でファイルを開きます");
                System.Diagnostics.Process.Start(filePath);

                if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellPosition))
                {
                    lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} を開きました。シート '{sheetName}' のセル {cellPosition} を確認してください。";
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                     lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} を開きました。シート '{sheetName}' を確認してください。";
                }
                else
                {
                    lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} を開きました。";
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"Excelファイルを開けませんでした: {ex.Message}";
                LogMessage($"エラー: {errorMsg}\n詳細: {ex.ToString()}", true);
                MessageBox.Show(errorMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Excel Interopを使用してファイルを開く実装（Late Binding方式）
        private bool OpenExcelWithInterop(string filePath, string sheetName, string cellPosition)
        {
            LogMessage($"OpenExcelWithInterop 開始 (Late Binding方式)");

            object excelApp = null;
            object workbook = null;

            try
            {
                // COMの状態を確認
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    LogMessage("警告: Excel.ApplicationのCOMタイプが見つかりません。Excel Interopが正しくインストールされていない可能性があります。");
                    return false;
                }

                // Excel アプリケーションを起動
                LogMessage("Excel アプリケーションのインスタンス作成中...");
                excelApp = Activator.CreateInstance(excelType);

                // バージョン情報を取得
                object version = excelType.InvokeMember("Version",
                    System.Reflection.BindingFlags.GetProperty,
                    null, excelApp, null);
                LogMessage($"Excel バージョン: {version}");

                // Visible プロパティを設定
                excelType.InvokeMember("Visible",
                    System.Reflection.BindingFlags.SetProperty,
                    null, excelApp, new object[] { true });

                // ファイルを開く
                LogMessage($"ファイルを開いています: {filePath}");

                // Workbooks コレクションを取得
                object workbooks = excelType.InvokeMember("Workbooks",
                    System.Reflection.BindingFlags.GetProperty,
                    null, excelApp, null);

                // Open メソッドを呼び出す
                Type workbooksType = workbooks.GetType();
                workbook = workbooksType.InvokeMember("Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null, workbooks, new object[] {
                        filePath,       // ファイルパス
                        Type.Missing,   // UpdateLinks
                        true,           // ReadOnly
                        Type.Missing,   // Format
                        Type.Missing,   // Password
                        Type.Missing,   // WriteResPassword
                        Type.Missing,   // IgnoreReadOnlyRecommended
                        Type.Missing,   // Origin
                        Type.Missing,   // Delimiter
                        Type.Missing,   // Editable
                        Type.Missing,   // Notify
                        Type.Missing,   // Converter
                        Type.Missing,   // AddToMru
                        Type.Missing,   // Local
                        Type.Missing    // CorruptLoad
                    });
                LogMessage("ワークブックを開きました");

                if (!string.IsNullOrEmpty(sheetName))
                {
                    // Sheets コレクションの取得
                    Type workbookType = workbook.GetType();
                    object sheets = workbookType.InvokeMember("Sheets",
                        System.Reflection.BindingFlags.GetProperty,
                        null, workbook, null);
                    Type sheetsType = sheets.GetType();

                    // シート名一覧をログに出力
                    LogMessage("利用可能なシート:");
                    object count = sheetsType.InvokeMember("Count",
                        System.Reflection.BindingFlags.GetProperty,
                        null, sheets, null);
                    for (int i = 1; i <= (int)count; i++)
                    {
                        object sheet = sheetsType.InvokeMember("Item",
                            System.Reflection.BindingFlags.GetProperty,
                            null, sheets, new object[] { i });
                        object name = sheet.GetType().InvokeMember("Name",
                            System.Reflection.BindingFlags.GetProperty,
                            null, sheet, null);
                        LogMessage($" - [{name}]");
                    }

                    // 指定されたシート名のシートを取得
                    LogMessage($"シート '{sheetName}' を検索中...");
                    object targetSheet = null;

                    try
                    {
                        targetSheet = sheetsType.InvokeMember("Item",
                            System.Reflection.BindingFlags.GetProperty,
                            null, sheets, new object[] { sheetName });
                        LogMessage($"シート '{sheetName}' が見つかりました");
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"シート取得エラー: {ex.Message}");
                        return !string.IsNullOrEmpty(cellPosition) && cellPosition != "図形内" && cellPosition != "図形内 (GF)" ? false : true;
                    }

                    // シートをアクティブにする
                    if (targetSheet != null)
                    {
                        LogMessage($"シート '{sheetName}' をアクティブ化します");
                        try
                        {
                            Type sheetType = targetSheet.GetType();
                            sheetType.InvokeMember("Activate",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, targetSheet, null);
                            LogMessage("シートのアクティブ化に成功しました");

                            // セル位置が指定されていれば選択
                            if (!string.IsNullOrEmpty(cellPosition) && cellPosition != "図形内" && cellPosition != "図形内 (GF)")
                            {
                                LogMessage($"セル {cellPosition} を選択します");
                                try
                                {
                                    // Rangeを取得
                                    object range = sheetType.InvokeMember("Range",
                                        System.Reflection.BindingFlags.GetProperty,
                                        null, targetSheet, new object[] { cellPosition });

                                    // セルを選択
                                    Type rangeType = range.GetType();
                                    rangeType.InvokeMember("Select",
                                        System.Reflection.BindingFlags.InvokeMethod,
                                        null, range, null);
                                    LogMessage("セルの選択に成功しました");
                                }
                                catch (Exception ex)
                                {
                                    LogMessage($"セル選択エラー: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage($"シートのアクティブ化エラー: {ex.Message}");
                        }
                    }
                    else
                    {
                        LogMessage($"警告: シート '{sheetName}' が見つかりませんでした");
                    }
                }

                // ステータス更新
                if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellPosition) && cellPosition != "図形内" && cellPosition != "図形内 (GF)")
                {
                    lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} のシート '{sheetName}' のセル {cellPosition} を開きました";
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                     lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} のシート '{sheetName}' を開きました";
                }
                else
                {
                    lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} を開きました";
                }
                LogMessage("Excel Interopの処理が完了しました (Late Binding方式)");
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"Excel Interopエラー (Late Binding): {ex.GetType().Name}: {ex.Message}");
                LogMessage($"スタックトレース: {ex.StackTrace}");
                return false;
            }
            finally
            {
                // リソースの解放
                try
                {
                    if (workbook != null)
                    {
                        LogMessage("COMオブジェクト (workbook) を解放します");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    }

                    if (excelApp != null)
                    {
                        LogMessage("COMオブジェクト (excelApp) を解放します");
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"COMオブジェクト解放エラー: {ex.Message}");
                }
            }
        }

        // GridViewのキーダウンイベントハンドラ
        private void GrdResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                // CTRL+A で全行選択
                LogMessage("Ctrl+A キーが押されました: 全行選択");
                grdResults.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                // CTRL+C でコピー
                LogMessage("Ctrl+C キーが押されました: 選択行をコピー");
                CopySelectedRowsToClipboard();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        // 選択行をクリップボードにコピーするメソッド
        private void CopySelectedRowsToClipboard()
        {
            LogMessage("選択行をクリップボードにコピー開始");
            if (grdResults.SelectedRows.Count > 0)
            {
                StringBuilder sb = new StringBuilder();

                // ヘッダー行を追加 (ヘッダーは通常エスケープ不要)
                for (int i = 0; i < grdResults.Columns.Count; i++)
                {
                    sb.Append(EscapeForClipboard(grdResults.Columns[i].HeaderText)); // ヘッダーも念のためエスケープ
                    sb.Append(i == grdResults.Columns.Count - 1 ? Environment.NewLine : "\t");
                }

                // 選択行を追加（選択順に処理）
                List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
                foreach (DataGridViewRow row in grdResults.SelectedRows)
                {
                    selectedRows.Add(row);
                }

                // インデックスでソート（上から下の順番になるように）
                selectedRows.Sort((x, y) => x.Index.CompareTo(y.Index));

                foreach (DataGridViewRow row in selectedRows)
                {
                    for (int i = 0; i < grdResults.Columns.Count; i++)
                    {
                        string cellValue = row.Cells[i].Value?.ToString() ?? "";
                        sb.Append(EscapeForClipboard(cellValue));
                        sb.Append(i == grdResults.Columns.Count - 1 ? Environment.NewLine : "\t");
                    }
                }

                try
                {
                    DataObject dataObject = new DataObject();
                    string textData = sb.ToString();
                    dataObject.SetData(DataFormats.UnicodeText, true, textData);
                    using (MemoryStream stream = new MemoryStream())
                    using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8)) 
                    {
                        writer.Write(textData);
                        writer.Flush();
                        stream.Position = 0;
                        dataObject.SetData("Csv", false, stream); 
                    }
                    Clipboard.SetDataObject(dataObject, true);
                    lblStatus.Text = $"{selectedRows.Count}行をクリップボードにコピーしました";
                    LogMessage($"{selectedRows.Count}行をクリップボードにコピーしました");
                }
                catch (Exception ex)
                {
                    LogMessage($"クリップボードへのコピーに失敗: {ex.Message}");
                    MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                LogMessage("コピー対象の行が選択されていません");
            }
        }

        // クリップボード用のテキストをエスケープするヘルパーメソッド
        private string EscapeForClipboard(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            bool needsQuoting = text.Contains("\r") || text.Contains("\n") || text.Contains("\t") || text.Contains("\"");
            if (needsQuoting)
            {
                string escapedText = text.Replace("\"", "\"\"");
                return "\"" + escapedText + "\"";
            }
            else
            {
                return text;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            LogMessage("アプリケーションを終了します");
            if (_isSearching && _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                LogMessage("実行中の検索をキャンセルします");
                _cancellationTokenSource.Cancel();
            }
            SaveSettings();
        }

        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;
            cmbFolderPath.Enabled = !isSearching;
            cmbKeyword.Enabled = !isSearching;
            cmbIgnoreKeywords.Enabled = !isSearching;
            chkRegex.Enabled = !isSearching;
            chkSearchShapes.Enabled = !isSearching;
            btnSelectFolder.Enabled = !isSearching;
            btnStartSearch.Enabled = !isSearching;
            btnCancelSearch.Enabled = isSearching;
            nudParallelism.Enabled = !isSearching;
            chkFirstHitOnly.Enabled = !isSearching;

            if (isSearching)
            {
                lblStatus.Text = "検索中...";
                LogMessage("検索を開始しました");
            }
            else
            {
                LogMessage("検索状態を終了しました");
            }
        }

        private async Task<List<SearchResult>> SearchExcelFilesAsync(
            string folderPath,
            string keyword,
            bool useRegex,
            Regex regex,
            List<string> ignoreKeywords,
            bool isRealTimeDisplay,
            bool searchShapes,
            CancellationToken cancellationToken)
        {
            LogMessage($"SearchExcelFilesAsync 開始: フォルダ={folderPath}");
            LogMessage($"検索開始: キーワード='{keyword}', 正規表現={useRegex}, 最初のヒットのみ={chkFirstHitOnly.Checked}, 図形内検索={searchShapes}");

            if (isRealTimeDisplay)
            {
                this.Invoke(new Action(() => grdResults.Rows.Clear()));
            }

            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.AllDirectories).ToArray();
            LogMessage($"{excelFiles.Length} 個のExcelファイル(.xlsx)が見つかりました");

            int totalFiles = excelFiles.Length;
            int processedFiles = 0;
            int maxParallelism = (int)nudParallelism.Value;
            bool firstHitOnly = chkFirstHitOnly.Checked;

            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = maxParallelism,
                CancellationToken = cancellationToken
            };

            var pendingResults = new System.Collections.Concurrent.ConcurrentQueue<SearchResult>();
            List<SearchResult> allResultsForThisSearch = new List<SearchResult>();

            using (var uiTimer = new System.Windows.Forms.Timer { Interval = 100 })
            {
                var uiUpdateEvent = new AutoResetEvent(false);

                uiTimer.Tick += (s, e) =>
                {
                    int batchSize = 0;
                    const int MaxUpdatesPerTick = 100;
                    while (batchSize < MaxUpdatesPerTick && pendingResults.TryDequeue(out SearchResult result))
                    {
                        allResultsForThisSearch.Add(result);
                        if (isRealTimeDisplay)
                        {
                            AddSearchResult(result);
                        }
                        batchSize++;
                    }
                    UpdateStatus($"処理中... {processedFiles}/{totalFiles} ファイル ({allResultsForThisSearch.Count} 件見つかりました)");
                    if (processedFiles >= totalFiles && pendingResults.IsEmpty)
                    {
                        uiUpdateEvent.Set();
                    }
                };

                uiTimer.Start();
                LogMessage("UIタイマーを開始しました。");

                try
                {
                    LogMessage($"並列処理を開始します (並列数: {maxParallelism})");
                    await Task.Run(() =>
                    {
                        Parallel.ForEach(excelFiles, parallelOptions, (filePath) =>
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (ignoreKeywords.Any(k => filePath.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                LogMessage($"無視キーワードが含まれるため処理をスキップ: {filePath}");
                                Interlocked.Increment(ref processedFiles);
                                return;
                            }
                            try
                            {
                                LogMessage($"ファイル処理開始: {filePath}");
                                string extension = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
                                if (extension == ".xlsx")
                                {
                                    List<SearchResult> fileResults = SearchInXlsxFileParallel(
                                        filePath, keyword, useRegex, regex, pendingResults,
                                        isRealTimeDisplay, firstHitOnly, searchShapes, cancellationToken);
                                    LogMessage($"ファイル処理完了: {filePath}, 見つかった結果(セル+図形): {fileResults.Count}件");
                                }
                                Interlocked.Increment(ref processedFiles);
                            }
                            catch (OperationCanceledException)
                            {
                                LogMessage($"ファイル処理がキャンセルされました: {filePath}");
                                throw;
                            }
                            catch (Exception ex)
                            {
                                LogMessage($"ファイル処理エラー: {filePath}, {ex.GetType().Name}: {ex.Message}");
                                Interlocked.Increment(ref processedFiles);
                            }
                        });
                    }, cancellationToken);

                    LogMessage("並列処理ループが完了しました。UI更新の完了を待機しています...");
                    while (processedFiles < totalFiles || !pendingResults.IsEmpty)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        await Task.Delay(50, cancellationToken);
                        LogMessage($"待機中: processedFiles={processedFiles}, totalFiles={totalFiles}, pendingResults.Count={pendingResults.Count}");
                    }
                    uiUpdateEvent.Set();
                    bool waitResult = uiUpdateEvent.WaitOne(TimeSpan.FromSeconds(5));
                    LogMessage($"UI更新待機結果: {waitResult}");
                }
                catch (OperationCanceledException)
                {
                    LogMessage("SearchExcelFilesAsync内のタスクがキャンセルされました。");
                    while (pendingResults.TryDequeue(out SearchResult result))
                    {
                        allResultsForThisSearch.Add(result);
                        if (isRealTimeDisplay) AddSearchResult(result);
                    }
                    throw;
                }
                catch (Exception ex)
                {
                    LogMessage($"SearchExcelFilesAsync内のタスクで予期せぬエラー: {ex.Message}");
                    while (pendingResults.TryDequeue(out SearchResult result))
                    {
                        allResultsForThisSearch.Add(result);
                        if (isRealTimeDisplay) AddSearchResult(result);
                    }
                    throw;
                }
                finally
                {
                    uiTimer.Stop();
                    LogMessage("UIタイマーを停止しました。");
                }
            }
            _searchResults.AddRange(allResultsForThisSearch);
            LogMessage($"SearchExcelFilesAsync 完了: {_searchResults.Count}件の結果");
            return _searchResults;
        }

        // XLSXファイル検索メソッド（セルと図形）
        private List<SearchResult> SearchInXlsxFileParallel(
            string filePath,
            string keyword,
            bool useRegex,
            Regex regex,
            System.Collections.Concurrent.ConcurrentQueue<SearchResult> pendingResults,
            bool isRealTimeDisplay,
            bool firstHitOnly,
            bool searchShapes,
            CancellationToken cancellationToken)
        {
            List<SearchResult> localResults = new List<SearchResult>();
            bool foundHitInFile = false;

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    if (workbookPart == null) return localResults;

                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (firstHitOnly && foundHitInFile) break;

                        string sheetName = GetSheetName(workbookPart, worksheetPart) ?? "不明なシート";

                        // 1. セル内のテキスト検索
                        if (worksheetPart.Worksheet != null)
                        {
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                            if (sheetData != null)
                            {
                                foreach (Row row in sheetData.Elements<Row>())
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    if (firstHitOnly && foundHitInFile) break;
                                    foreach (Cell cell in row.Elements<Cell>())
                                    {
                                        cancellationToken.ThrowIfCancellationRequested();
                                        if (firstHitOnly && foundHitInFile) break;
                                        string cellValue = GetCellValue(cell, sharedStringTable);
                                        if (!string.IsNullOrEmpty(cellValue))
                                        {
                                            bool isMatch = useRegex && regex != null ?
                                                           regex.IsMatch(cellValue) :
                                                           cellValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                                            if (isMatch)
                                            {
                                                SearchResult result = new SearchResult
                                                {
                                                    FilePath = filePath, SheetName = sheetName,
                                                    CellPosition = GetCellReference(cell), CellValue = cellValue
                                                };
                                                pendingResults.Enqueue(result);
                                                localResults.Add(result);
                                                foundHitInFile = true;
                                                LogMessage($"セル内一致: {filePath} - {sheetName} - {result.CellPosition} - '{TruncateString(result.CellValue)}'");
                                                if (firstHitOnly) break;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // 2. 図形内のテキスト検索
                        if (searchShapes)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (firstHitOnly && foundHitInFile && !localResults.Any(lr => lr.CellPosition.StartsWith("図形内"))) break;

                            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
                            if (drawingsPart != null && drawingsPart.WorksheetDrawing != null)
                            {
                                foreach (var twoCellAnchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    if (firstHitOnly && foundHitInFile && !localResults.Any(lr => lr.CellPosition.StartsWith("図形内"))) break;

                                    var shape = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault();
                                    if (shape != null && shape.TextBody != null)
                                    {
                                        string shapeText = GetTextFromShapeTextBody(shape.TextBody); // <--- 修正された呼び出し
                                        if (!string.IsNullOrEmpty(shapeText))
                                        {
                                            bool isMatch = useRegex && regex != null ? regex.IsMatch(shapeText) : shapeText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                                            if (isMatch)
                                            {
                                                SearchResult result = new SearchResult
                                                {
                                                    FilePath = filePath, SheetName = sheetName,
                                                    CellPosition = "図形内", CellValue = TruncateString(shapeText)
                                                };
                                                pendingResults.Enqueue(result);
                                                localResults.Add(result);
                                                foundHitInFile = true;
                                                LogMessage($"図形内一致 (Shape): {filePath} - {sheetName} - '{result.CellValue}'");
                                                if (firstHitOnly) break;
                                            }
                                        }
                                    }

                                    var graphicFrame = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault();
                                    if (graphicFrame != null)
                                    {
                                        string frameText = GetTextFromGraphicFrame(graphicFrame); // <--- 修正された呼び出し
                                        if (!string.IsNullOrEmpty(frameText))
                                        {
                                            bool isMatch = useRegex && regex != null ? regex.IsMatch(frameText) : frameText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                                            if (isMatch)
                                            {
                                                SearchResult result = new SearchResult
                                                {
                                                    FilePath = filePath, SheetName = sheetName,
                                                    CellPosition = "図形内 (GF)", CellValue = TruncateString(frameText)
                                                };
                                                pendingResults.Enqueue(result);
                                                localResults.Add(result);
                                                foundHitInFile = true;
                                                LogMessage($"図形内一致 (GraphicFrame): {filePath} - {sheetName} - '{result.CellValue}'");
                                                if (firstHitOnly) break;
                                            }
                                        }
                                    }
                                    if (firstHitOnly && foundHitInFile) break;
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LogMessage($"ファイル処理がキャンセルされました (SearchInXlsxFileParallel): {filePath}");
                throw;
            }
            catch (Exception ex)
            {
                LogMessage($"Excel処理エラー (SearchInXlsxFileParallel): {filePath}, {ex.GetType().Name}: {ex.Message}");
            }
            return localResults;
        }

        private string GetSheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            string sheetId = workbookPart.GetIdOfPart(worksheetPart);
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Id?.Value == sheetId);
            return sheet?.Name?.Value;
        }

        private string TruncateString(string value, int maxLength = 255)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength) + "...";
        }


        // 修正点：ジェネリックヘルパーメソッドの導入
        private string ExtractTextFromTextBodyElement<T>(T textBodyElement) where T : DocumentFormat.OpenXml.OpenXmlCompositeElement
        {
            StringBuilder sb = new StringBuilder();
            if (textBodyElement == null) return string.Empty;

            foreach (var paragraph in textBodyElement.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                {
                    var textElement = run.Elements<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                    if (textElement != null)
                    {
                        sb.Append(textElement.InnerText);
                    }
                }
                // 段落ごとに改行を挿入（ただし、最後の空の段落は除く）
                if (sb.Length > 0 && !sb.ToString().EndsWith(Environment.NewLine) && paragraph.HasChildren)
                {
                     sb.AppendLine();
                }
            }
            // 末尾の余分な改行を削除
            return sb.ToString().TrimEnd(Environment.NewLine.ToCharArray());
        }

        // ShapeのTextBodyからテキストを抽出するヘルパーメソッド
        // 引数の型を DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody に変更
        private string GetTextFromShapeTextBody(DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody textBody)
        {
            return ExtractTextFromTextBodyElement(textBody);
        }

        // GraphicFrameからテキストを抽出するヘルパーメソッド
        private string GetTextFromGraphicFrame(DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame)
        {
            StringBuilder sb = new StringBuilder();
            var graphicData = graphicFrame.Graphic?.GraphicData;
            if (graphicData != null)
            {
                // GraphicData 直下の TextBody (D.TextBody) を探す
                var directTextBodies = graphicData.Elements<DocumentFormat.OpenXml.Drawing.TextBody>();
                foreach(var textBody in directTextBodies)
                {
                    sb.Append(ExtractTextFromTextBodyElement(textBody));
                }

                // さらに深い階層の TextBody (D.TextBody) も探す (例: Chart内など)
                // ただし、無限ループや意図しない要素を拾わないように注意が必要
                // ここでは Descendants を使って簡易的に取得
                var descendantTextBodies = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>();
                foreach(var textBody in descendantTextBodies.Except(directTextBodies)) // 重複を避ける
                {
                     if (sb.Length > 0 && textBody.HasChildren) sb.AppendLine(); // 複数のTextBodyが見つかった場合の区切り
                    sb.Append(ExtractTextFromTextBodyElement(textBody));
                }
            }
            return sb.ToString().Trim();
        }


        // セルの値を取得するヘルパーメソッド
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            string cellValueStr = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
            {
                if (int.TryParse(cellValueStr, out int ssid) && ssid >= 0 && ssid < sharedStringTable.ChildElements.Count)
                {
                    SharedStringItem ssi = sharedStringTable.ChildElements[ssid] as SharedStringItem;
                    if (ssi != null)
                    {
                        // Text 要素の値を連結する
                        return string.Concat(ssi.Elements<Text>().Select(t => t.Text));
                    }
                }
                return string.Empty; // 共有文字列が見つからない場合
            }
            return cellValueStr;
        }

        // セル参照（例：A1）を取得するメソッド
        private string GetCellReference(Cell cell)
        {
            return cell.CellReference?.Value ?? string.Empty;
        }

        private void UpdateStatus(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() => lblStatus.Text = message));
            }
            else
            {
                lblStatus.Text = message;
            }
        }
    }
}