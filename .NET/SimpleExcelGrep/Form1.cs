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
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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
            public string CellPosition { get; set; }
            public string CellValue { get; set; }
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
                        regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled);
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
                    chkSearchShapes.Checked,
                    _cancellationTokenSource.Token);

                // リアルタイム表示がOFFの場合または検索が途中でキャンセルされた場合に、
                // 最終的な結果をまとめて表示
                if (!isRealTimeDisplay || _cancellationTokenSource.IsCancellationRequested)
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

                OpenExcelFile(filePath, sheetName, cellPosition);
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
                
                lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} を開きました。シート '{sheetName}' のセル {cellPosition} を確認してください。";
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
                        return true; // シートが見つからなくてもファイルは開けているのでtrueを返す
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
                            if (!string.IsNullOrEmpty(cellPosition))
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
                lblStatus.Text = $"{System.IO.Path.GetFileName(filePath)} のシート '{sheetName}' のセル {cellPosition} を開きました";
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

                        // ★★★ 修正箇所: セル値をエスケープする ★★★
                        sb.Append(EscapeForClipboard(cellValue));
                        // ★★★★★★★★★★★★★★★★★★★★★★

                        sb.Append(i == grdResults.Columns.Count - 1 ? Environment.NewLine : "\t");
                    }
                }

                try
                {
                    // DataObjectを使用して、テキスト形式とCSV形式の両方で設定する
                    // これにより、Excelはより適切にデータを解釈できる可能性が高まる
                    DataObject dataObject = new DataObject();
                    string textData = sb.ToString();

                    // 通常のテキスト形式
                    dataObject.SetData(DataFormats.UnicodeText, true, textData);
                    // CSV形式 (Excelが優先的に解釈する可能性)
                    // メモリ ストリームに UTF-8 (BOM 付き) で書き込む
                    using (MemoryStream stream = new MemoryStream())
                    using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8)) // BOM付きUTF-8
                    {
                        writer.Write(textData);
                        writer.Flush();
                        stream.Position = 0;
                        dataObject.SetData("Csv", false, stream); // "Csv" フォーマット
                    }

                    Clipboard.SetDataObject(dataObject, true);

                    // コピー成功を通知（オプション）
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

            // 改行、タブ、ダブルクォーテーションが含まれる場合はエスケープが必要
            bool needsQuoting = text.Contains("\r") || text.Contains("\n") || text.Contains("\t") || text.Contains("\"");

            if (needsQuoting)
            {
                // 1. 内部のダブルクォーテーションを二重にする ("" に置換)
                string escapedText = text.Replace("\"", "\"\"");
                // 2. 全体をダブルクォーテーションで囲む
                return "\"" + escapedText + "\"";
            }
            else
            {
                // エスケープ不要な場合はそのまま返す
                return text;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            LogMessage("アプリケーションを終了します");
            // 検索中の場合、キャンセルする
            if (_isSearching && _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                LogMessage("実行中の検索をキャンセルします");
                _cancellationTokenSource.Cancel();
                // キャンセル完了を少し待つ（非同期処理のため即時終了しない場合がある）
                // Task.Delay(500).Wait(); // 必要に応じて調整
            }

            // 設定を保存
            SaveSettings();
        }

        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;

            // 検索中はUIの一部を無効化
            cmbFolderPath.Enabled = !isSearching;
            cmbKeyword.Enabled = !isSearching;
            cmbIgnoreKeywords.Enabled = !isSearching;
            chkRegex.Enabled = !isSearching;
            btnSelectFolder.Enabled = !isSearching;
            btnStartSearch.Enabled = !isSearching;
            btnCancelSearch.Enabled = isSearching;
            
            // リアルタイム表示チェックボックスは検索中も有効のまま
            // chkRealTimeDisplay.Enabled は変更しない
            
            // 並列数設定も検索中は無効化
            nudParallelism.Enabled = !isSearching;
            
            // 最初のヒットのみチェックボックスも検索中は無効化
            chkFirstHitOnly.Enabled = !isSearching;

            // 検索中はステータスを更新
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
    
            // 結果グリッドをクリア
            if (isRealTimeDisplay)
            {
                this.Invoke(new Action(() => grdResults.Rows.Clear()));
            }

            // Excelファイルの一覧を取得
            LogMessage("Excelファイルを検索しています...");
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xls*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || 
                            f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            LogMessage($"{excelFiles.Length} 個のExcelファイルが見つかりました");

            // 進捗状況を表示するための変数
            int totalFiles = excelFiles.Length;
            int processedFiles = 0;
    
            // 現在の並列処理数を取得
            int maxParallelism = (int)nudParallelism.Value;
    
            // 最初のヒットのみ設定を取得
            bool firstHitOnly = chkFirstHitOnly.Checked;
    
            // 並列処理のオプション
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = maxParallelism,
                CancellationToken = cancellationToken
            };
    
            // UI更新のためのキュー
            var pendingResults = new System.Collections.Concurrent.ConcurrentQueue<SearchResult>();
    
            // 全体の結果を格納するリスト（デバッグ用）
            List<SearchResult> allResults = new List<SearchResult>();
    
            // UIの更新タイマー
            using (var uiTimer = new System.Windows.Forms.Timer { Interval = 100 })
            {
                var uiUpdateEvent = new AutoResetEvent(false);
        
                uiTimer.Tick += (s, e) =>
                {
                    // 現在のキューサイズをログに記録
                    LogMessage($"UI Timer tick: キューサイズ={pendingResults.Count}, 処理済みファイル={processedFiles}/{totalFiles}");
    
                    // キューが空でない場合、内容を確認する
                    if (pendingResults.Count > 0)
                    {
                        LogMessage($"キューにデータがあります: {pendingResults.Count}件");
        
                        try
                        {
                            int batchSize = 0;
                            const int MaxUpdatesPerTick = 50;
            
                            // キューからデータを取り出して処理
                            while (batchSize < MaxUpdatesPerTick && pendingResults.TryDequeue(out SearchResult result))
                            {
                                LogMessage($"キューから取得: シート={result.SheetName}, セル={result.CellPosition}, 値={result.CellValue}");
                
                                // 結果リストに追加
                                _searchResults.Add(result);
                                LogMessage($"_searchResultsに追加: 現在のサイズ={_searchResults.Count}");
                
                                // リアルタイム表示が有効な場合はUIに追加
                                if (isRealTimeDisplay)
                                {
                                    AddSearchResult(result);
                                    LogMessage($"グリッドに追加: シート={result.SheetName}, セル={result.CellPosition}");
                                }
                
                                batchSize++;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage($"UI更新中にエラー発生: {ex.Message}");
                        }
                    }
    
                    // 進捗状況を更新
                    UpdateStatus($"処理中... {processedFiles}/{totalFiles} ファイル ({_searchResults.Count} 件見つかりました)");
    
                    // すべての処理が完了したかを確認
                    if (processedFiles >= totalFiles)
                    {
                        LogMessage($"すべてのファイル処理完了。キューサイズ={pendingResults.Count}, 結果件数={_searchResults.Count}");
        
                        if (pendingResults.IsEmpty)
                        {
                            LogMessage("キューが空になりました。UI更新を完了します。");
                            uiUpdateEvent.Set();
                        }
                    }
                };
        
                uiTimer.Start();

                try
                {
                    // 並列でファイルを処理
                    LogMessage($"並列処理を開始します (並列数: {maxParallelism})");
                    await Task.Run(() =>
                    {
                        Parallel.ForEach(excelFiles, parallelOptions, (filePath) =>
                        {
                            // キャンセルされた場合は処理を中止
                            cancellationToken.ThrowIfCancellationRequested();
                    
                            // 無視キーワードが含まれている場合はスキップ
                            if (ignoreKeywords.Any(k => filePath.Contains(k)))
                            {
                                LogMessage($"無視キーワードが含まれるため処理をスキップ: {filePath}");
                                Interlocked.Increment(ref processedFiles);
                                return;
                            }

                            try
                            {
                                LogMessage($"ファイル処理: {filePath}");
                                LogMessage($"ファイル処理開始: {filePath}");
                                // ファイル拡張子によって処理を分ける
                                string extension = System.IO.Path.GetExtension(filePath).ToLower();
                                List<SearchResult> fileResults = new List<SearchResult>();

                                if (extension == ".xlsx")
                                {
                                    // .xlsx ファイルを処理
                                    fileResults = SearchInXlsxFileParallel(
                                        filePath, keyword, useRegex, regex, pendingResults, 
                                        isRealTimeDisplay, firstHitOnly, searchShapes, cancellationToken);
                            
                                    LogMessage($"ファイル処理完了: {filePath}, 見つかった結果: {fileResults.Count}件");
                            
                                    // デバッグ用に全結果を保存
                                    lock(allResults)
                                    {
                                        allResults.AddRange(fileResults);
                                    }
                                }
                                else if (extension == ".xls")
                                {
                                    // .xls ファイルはサポート外として処理
                                    LogMessage($"未サポートの形式: {filePath}");
                                    this.Invoke(new Action(() => 
                                        UpdateStatus($"注: .xls形式は現在サポートされていません: {filePath}")));
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
                                // ファイル処理中のエラーをログに記録
                                LogMessage($"ファイル処理エラー: {filePath}, {ex.Message}");
                                Interlocked.Increment(ref processedFiles);
                            }
                        });
                    }, cancellationToken);
            
                    // すべてのUI更新が完了するまで待機
                    LogMessage("すべてのファイル処理が完了しました。UI更新の完了を待機中...");
                    LogMessage($"現在の状態: 処理済みファイル={processedFiles}/{totalFiles}, キューサイズ={pendingResults.Count}, 結果件数={_searchResults.Count}");
            
                    bool waitResult = uiUpdateEvent.WaitOne(5000); // 最大5秒待機
                    LogMessage($"待機結果: {waitResult}, 最終的な結果件数={_searchResults.Count}");
            
                    // 最終確認：結果が見つかっているのに_searchResultsが空の場合は、直接コピーする
                    // SearchExcelFilesAsyncメソッド内のTask.Run完了後に追加
                    LogMessage($"並列処理完了後の結果数: {allResults.Count}件");
                    _searchResults.AddRange(allResults);
                    if (isRealTimeDisplay || _searchResults.Count > 0)
                    {
                        this.Invoke(new Action(() => DisplaySearchResults(_searchResults)));
                        LogMessage($"並列処理完了後、結果を直接表示しました: {_searchResults.Count}件");
                    }
                }
                catch (OperationCanceledException)
                {
                    LogMessage("検索操作がキャンセルされました");
                    throw;
                }
                catch (Exception ex)
                {
                    LogMessage($"検索中に例外が発生しました: {ex.Message}");
                    throw;
                }
                finally
                {
                    uiTimer.Stop();
                    LogMessage("UIタイマーを停止しました");
                }
            }

            LogMessage($"SearchExcelFilesAsync 完了: {_searchResults.Count}件の結果");
            return _searchResults;
        }

        // 並列処理用に修正したXLSXファイル検索メソッド
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
            List<SearchResult> results = new List<SearchResult>();
            bool foundHit = false; // ヒット検出フラグ

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
                    SharedStringTable sharedStringTable = sharedStringTablePart?.SharedStringTable;

                    // ワークシートの一覧を取得
                    Sheets sheets = workbookPart.Workbook.Sheets;
                    
                    // 各シートを処理
                    foreach (Sheet sheet in sheets.Elements<Sheet>())
                    {
                        // キャンセル処理
                        cancellationToken.ThrowIfCancellationRequested();
                        
                        // 既にヒットがあり、最初のヒットのみモードの場合は処理を終了
                        if (firstHitOnly && foundHit)
                            break;

                        // シートIDを取得
                        string relationshipId = sheet.Id.Value;
                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                        // 各行を処理
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            // キャンセル処理
                            cancellationToken.ThrowIfCancellationRequested();
                            
                            // 既にヒットがあり、最初のヒットのみモードの場合は処理を終了
                            if (firstHitOnly && foundHit)
                                break;

                            // 各セルを処理
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                // キャンセル処理
                                cancellationToken.ThrowIfCancellationRequested();
                                
                                // 既にヒットがあり、最初のヒットのみモードの場合は処理を終了
                                if (firstHitOnly && foundHit)
                                    break;

                                // セルの値を取得
                                string cellValue = GetCellValue(cell, sharedStringTable);

                                Debug.WriteLine($"セル {GetCellReference(cell)} の値: '{cellValue}'");
                                
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    bool isMatch;
                                    
                                    if (useRegex && regex != null)
                                    {
                                        isMatch = regex.IsMatch(cellValue);
                                        Debug.WriteLine($"正規表現マッチ: {isMatch}, パターン='{cmbKeyword.Text}', 値='{cellValue}'");
                                    }
                                    else
                                    {
                                        // 大文字小文字を区別しない比較に変更
                                        isMatch = cellValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
                                        Debug.WriteLine($"文字列比較: {isMatch}, キーワード='{keyword}', 値='{cellValue}'");
                                    }

                                    // マッチするセルを見つけた場合
                                    if (isMatch)
                                    {
                                        SearchResult result = new SearchResult
                                        {
                                            FilePath = filePath,
                                            SheetName = sheet.Name,
                                            CellPosition = GetCellReference(cell),
                                            CellValue = cellValue
                                        };
    
                                        results.Add(result);
    
                                        // 明示的にログを出力
                                        LogMessage($"見つかった結果: シート={sheet.Name}, セル={GetCellReference(cell)}, 値={cellValue}");
    
                                        // キューに追加
                                        pendingResults.Enqueue(result);
                                        LogMessage($"pendingResultsキューに追加: 現在のキューサイズ={pendingResults.Count}");
    
                                        foundHit = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                // エラーをログに記録
                LogMessage($"Excel処理エラー: {filePath}, {ex.Message}");
            }

            return results;
        }

        // セルの値を取得するヘルパーメソッド
        private string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {
            if (cell == null)
                return string.Empty;

            string cellValue = cell.InnerText;
            Debug.WriteLine($"GetCellValue: セル参照={cell.CellReference?.Value ?? "不明"}, DataType={cell.DataType?.Value.ToString() ?? "null"}, InnerText='{cellValue}'");

            // セルの値が共有文字列の場合
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString && sharedStringTable != null)
            {
                if (int.TryParse(cellValue, out int ssid) && ssid >= 0 && ssid < sharedStringTable.Count())
                {
                    SharedStringItem ssi = sharedStringTable.Elements<SharedStringItem>().ElementAt(ssid);
            
                    // TextElementとInnerTextの両方を試す
                    string textFromText = ssi.Text?.Text ?? "";
                    string textFromInnerText = ssi.InnerText ?? "";
            
                    Debug.WriteLine($"  共有文字列: ssid={ssid}, Text='{textFromText}', InnerText='{textFromInnerText}'");
            
                    if (!string.IsNullOrEmpty(textFromText))
                        return textFromText;
                    else if (!string.IsNullOrEmpty(textFromInnerText))
                        return textFromInnerText;
                    else if (!string.IsNullOrEmpty(ssi.InnerXml))
                        return ssi.InnerXml;
                }
                else
                {
                    Debug.WriteLine($"  無効な共有文字列ID: {cellValue}");
                }
            }
    
            // それ以外の場合はそのまま返す
            return cellValue;
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

        // 図形内テキスト取得の実験 >>
        private void button1_Click ( object sender, EventArgs e )
        {
            try
            {
                string filePath = @"C:\_git\Public\.NET\SimpleExcelGrep\testdata\test.xlsx"; // 対象のExcelファイルパス
                string searchText = "HOGE"; // 検索する文字列

                ExcelShapeTextFinder.FindTextInShapes(filePath, searchText);
            }
            catch ( Exception ex )
            {
                MessageBox.Show( $"エラー: {ex.Message}" );
            }
            finally
            {
            }
        }
        // 図形内テキスト取得の実験 <<
    }

    // 図形内テキスト取得の実験 >>
    public class ExcelShapeTextFinder
    {
        public static void FindTextInShapes(string filePath, string searchText)
        {
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart == null) return;

                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                {
                    DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
                    if (drawingsPart != null)
                    {
                        // Xdr (TwoCellAnchor) や Wps (Shape) など、図形の種類によって構造が異なる場合があります。
                        // ここでは一般的なShape内のテキストを想定します。

                        // TwoCellAnchor は図形の位置などを定義します。
                        foreach (var twoCellAnchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                        {
                            // Shape要素を取得
                            var shape = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault();
                            if (shape != null && shape.TextBody != null)
                            {
                                foreach (var paragraph in shape.TextBody.Elements<Paragraph>())
                                {
                                    foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                                    {
                                        var text = run.Text;
                                        if (text != null && text.InnerText.Contains(searchText))
                                        {
                                            // ここで一致した図形に対する処理を記述
                                            // 例: 図形のIDや位置、テキストなどを出力
                                            System.Console.WriteLine($"Found '{searchText}' in a shape on worksheet: {worksheetPart.Uri}");
                                            // 必要であれば、図形の詳細情報（位置など）も取得できます。
                                            // 例: fromRow = twoCellAnchor.FromMarker.RowId.InnerText;
                                            //     fromCol = twoCellAnchor.FromMarker.ColumnId.InnerText;
                                        }
                                    }
                                }
                            }

                            //// GraphicFrame 内のテキストも考慮する場合 (例: SmartArt やグラフ内のテキストボックスなど)
                            //var graphicFrame = twoCellAnchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault();
                            //if (graphicFrame != null)
                            //{
                            //    // GraphicFrame内のテキスト構造はさらに複雑になることがあります。
                            //    // a:graphicData/a:txBody/a:p/a:r/a:t のような階層を辿る必要があります。
                            //    var graphicData = graphicFrame.Graphic?.GraphicData;
                            //    if (graphicData != null)
                            //    {
                            //        // ここでは簡略化のため、テキストボディ内の段落とランを直接探すようなイメージです。
                            //        // 実際には、より詳細な要素の探索が必要になる場合があります。
                            //        foreach (var p in graphicData.Descendants<Paragraph>())
                            //        {
                            //            foreach (var r in p.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                            //            {
                            //                var t = r.Text;
                            //                if (t != null && t.InnerText.Contains(searchText))
                            //                {
                            //                    System.Console.WriteLine($"Found '{searchText}' in a graphic frame on worksheet: {worksheetPart.Uri}");
                            //                }
                            //            }
                            //        }
                            //    }
                            //}
                        }
                    }
                }
            }
        }
    }
    // 図形内テキスト取得の実験 <<
}