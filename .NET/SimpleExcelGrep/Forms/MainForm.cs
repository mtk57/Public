using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SimpleExcelGrep.Models;
using SimpleExcelGrep.Services;
using SimpleExcelGrep.Utilities;

namespace SimpleExcelGrep.Forms
{
    public partial class MainForm : Form
    {
        // サービスクラス
        private readonly LogService _logService;
        private readonly SettingsService _settingsService;
        private readonly ExcelSearchService _excelSearchService;
        private readonly ExcelInteropService _excelInteropService;

        // 状態変数
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isSearching = false;
        private bool _isLoading = false;
        private List<SearchResult> _searchResults = new List<SearchResult>();

        private System.Windows.Forms.Timer _uiTimer;

        public MainForm()
        {
            InitializeComponent();
    
            // サービスの初期化
            _logService = new LogService(lblStatus);
            _settingsService = new SettingsService(_logService);
            _excelSearchService = new ExcelSearchService(_logService);
            _excelInteropService = new ExcelInteropService(_logService);
            _uiTimer = null;
    
            // イベントハンドラの登録
            RegisterEventHandlers();
    
            // UIの初期設定
            InitializeUI();
        }

        /// <summary>
        /// イベントハンドラを登録
        /// </summary>
        private void RegisterEventHandlers()
        {
            this.FormClosing += MainForm_FormClosing;
            this.Load += MainForm_Load;
            btnSelectFolder.Click += BtnSelectFolder_Click;
            btnStartSearch.Click += BtnStartSearch_Click;
            btnCancelSearch.Click += BtnCancelSearch_Click;
            grdResults.DoubleClick += GrdResults_DoubleClick;
            grdResults.KeyDown += GrdResults_KeyDown;
            cmbKeyword.KeyDown += CmbKeyword_KeyDown;
            
            // チェックボックスとNumericUpDownのイベント
            chkRealTimeDisplay.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            nudParallelism.ValueChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            chkFirstHitOnly.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            chkSearchShapes.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
        }

        /// <summary>
        /// UIの初期設定
        /// </summary>
        private void InitializeUI()
        {
            // 複数行選択を可能にする
            grdResults.MultiSelect = true;
            grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            // コンテキストメニューの追加
            ContextMenuStrip contextMenu = new ContextMenuStrip();
            
            ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("コピー");
            copyMenuItem.Click += (s, e) => CopySelectedRowsToClipboard();
            contextMenu.Items.Add(copyMenuItem);

            ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("すべて選択");
            selectAllMenuItem.Click += (s, e) => grdResults.SelectAll();
            contextMenu.Items.Add(selectAllMenuItem);

            grdResults.ContextMenuStrip = contextMenu;
        }

        /// <summary>
        /// フォーム読み込み時の処理
        /// </summary>
        private void MainForm_Load(object sender, EventArgs e)
        {
            _logService.LogMessage("アプリケーション起動");
            _logService.LogEnvironmentInfo();
            LoadSettings();
        }

        /// <summary>
        /// 設定をUIに読み込む
        /// </summary>
        private void LoadSettings()
        {
            _isLoading = true;

            try
            {
                Settings settings = _settingsService.LoadSettings();

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
                chkRealTimeDisplay.Checked = settings.RealTimeDisplay;
                nudParallelism.Value = Math.Min(Math.Max(settings.MaxParallelism, 1), 32);
                chkFirstHitOnly.Checked = settings.FirstHitOnly;
                chkSearchShapes.Checked = settings.SearchShapes;
            }
            finally
            {
                _isLoading = false;
            }
        }

        /// <summary>
        /// 現在の設定を保存
        /// </summary>
        private void SaveCurrentSettings()
        {
            if (_isLoading) return;

            Settings settings = new Settings
            {
                FolderPath = cmbFolderPath.Text,
                SearchKeyword = cmbKeyword.Text,
                UseRegex = chkRegex.Checked,
                IgnoreKeywords = cmbIgnoreKeywords.Text,
                FolderPathHistory = _settingsService.CreateHistoryListFromComboBox(cmbFolderPath),
                SearchKeywordHistory = _settingsService.CreateHistoryListFromComboBox(cmbKeyword),
                IgnoreKeywordsHistory = _settingsService.CreateHistoryListFromComboBox(cmbIgnoreKeywords),
                RealTimeDisplay = chkRealTimeDisplay.Checked,
                MaxParallelism = (int)nudParallelism.Value,
                FirstHitOnly = chkFirstHitOnly.Checked,
                SearchShapes = chkSearchShapes.Checked
            };

            if (!_settingsService.SaveSettings(settings))
            {
                MessageBox.Show("設定の保存に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// フォルダ選択ボタンクリック時の処理
        /// </summary>
        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("フォルダ選択ボタンがクリックされました");
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
                    _logService.LogMessage($"選択されたフォルダ: {dialog.SelectedPath}");
                    _settingsService.AddToComboBoxHistory(cmbFolderPath, dialog.SelectedPath);
                    SaveCurrentSettings();
                }
                else
                {
                    _logService.LogMessage("フォルダ選択はキャンセルされました");
                }
            }
        }

        /// <summary>
        /// 検索キーワードでEnterキー押下時の処理
        /// </summary>
        private void CmbKeyword_KeyDown(object sender, KeyEventArgs e)
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

        /// <summary>
        /// 検索開始ボタンクリック時の処理
        /// </summary>
        private async void BtnStartSearch_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("検索開始ボタンがクリックされました");

            // 入力検証
            if (!ValidateSearchInputs()) return;

            // 検索キーワードと無視キーワードを履歴に追加
            _settingsService.AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);
            _settingsService.AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);
            SaveCurrentSettings();

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

                _logService.LogMessage($"無視キーワードリスト: {string.Join(", ", ignoreKeywords)}");

                // 正規表現オブジェクト
                Regex regex = null;
                if (chkRegex.Checked)
                {
                    try
                    {
                        _logService.LogMessage($"正規表現を使用: {cmbKeyword.Text}");
                        regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                    }
                    catch (Exception ex)
                    {
                        _logService.LogMessage($"正規表現エラー: {ex.Message}");
                        MessageBox.Show($"正規表現が無効です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetSearchingState(false);
                        return;
                    }
                }

                // 検索処理を実行
                bool isRealTimeDisplay = chkRealTimeDisplay.Checked;
                bool firstHitOnly = chkFirstHitOnly.Checked;
                bool searchShapes = chkSearchShapes.Checked;
                int maxParallelism = (int)nudParallelism.Value;
        
                _logService.LogMessage($"リアルタイム表示: {isRealTimeDisplay}, 最初のヒットのみ: {firstHitOnly}, 図形内検索: {searchShapes}, 並列数: {maxParallelism}");

                // 検索結果をキューで受け取るための準備
                var pendingResults = new ConcurrentQueue<SearchResult>();

                // タイマーを作成して開始（クラスフィールドを使用）
                _uiTimer = new System.Windows.Forms.Timer { Interval = 100 };
                var uiUpdateEvent = new AutoResetEvent(false);

                // UI更新のためのタイマー処理を設定
                StartResultUpdateTimer(_uiTimer, pendingResults, isRealTimeDisplay, uiUpdateEvent);
                _logService.LogMessage($"StartResultUpdateTimer呼び出し後: タイマーの状態={_uiTimer.Enabled}");

                try
                {
                    // 検索を実行
                    _searchResults = await _excelSearchService.SearchExcelFilesAsync(
                        cmbFolderPath.Text,
                        cmbKeyword.Text,
                        chkRegex.Checked,
                        regex,
                        ignoreKeywords,
                        isRealTimeDisplay,
                        searchShapes,
                        firstHitOnly,
                        maxParallelism,
                        pendingResults,
                        UpdateStatus,
                        _cancellationTokenSource.Token);
    
                    // リアルタイム表示でも非リアルタイム表示でも、最終結果を確実に表示する
                    _logService.LogMessage($"最終結果をまとめて表示: {_searchResults.Count}件");
                    DisplaySearchResults(_searchResults);
    
                    // 一度実行されていることを確認するために追加の処理実行
                    _logService.LogMessage($"pendingResultsキューの残り: {pendingResults.Count}件");
                    List<SearchResult> remainingResults = new List<SearchResult>();
                    while (pendingResults.TryDequeue(out SearchResult result))
                    {
                        remainingResults.Add(result);
                    }
                    if (remainingResults.Count > 0)
                    {
                        _logService.LogMessage($"キューから{remainingResults.Count}件の未処理結果を追加");
                        foreach (var result in remainingResults)
                        {
                            if (!_searchResults.Contains(result))
                            {
                                _searchResults.Add(result);
                            }
                        }
                        // 結果を再表示
                        DisplaySearchResults(_searchResults);
                    }
    
                    UpdateStatus($"検索完了: {_searchResults.Count} 件見つかりました");
                }
                catch (OperationCanceledException)
                {
                    UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました");

                    // キャンセル時は現在までの結果を表示
                    if (!isRealTimeDisplay)
                    {
                        DisplaySearchResults(_searchResults);
                    }
                }
                catch (Exception ex)
                {
                    _logService.LogMessage($"検索中にエラーが発生しました: {ex.Message}");
                    MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus("エラーが発生しました");

                    // エラー時は現在までの結果を表示
                    if (!isRealTimeDisplay)
                    {
                        DisplaySearchResults(_searchResults);
                    }
                }
                finally
                {
                    // タイマーを停止して破棄
                    if (_uiTimer != null)
                    {
                        _uiTimer.Stop();
                        _uiTimer.Dispose();
                        _uiTimer = null;
                        _logService.LogMessage("UIタイマーを停止しました。");
                    }
                }
            }
            finally
            {
                // UIを通常の状態に戻す
                SetSearchingState(false);
            }
        }

        /// <summary>
        /// 結果更新用タイマーを開始
        /// </summary>
        private void StartResultUpdateTimer(System.Windows.Forms.Timer timer, ConcurrentQueue<SearchResult> pendingResults, bool isRealTimeDisplay, AutoResetEvent uiUpdateEvent)
        {
            timer.Tick += (s, e) =>
            {
                int batchSize = 0;
                const int MaxUpdatesPerTick = 100;
                List<SearchResult> tempResults = new List<SearchResult>();
        
                while (batchSize < MaxUpdatesPerTick && pendingResults.TryDequeue(out SearchResult result))
                {
                    _searchResults.Add(result);
                    tempResults.Add(result);
                    batchSize++;
                }
        
                if (isRealTimeDisplay && tempResults.Count > 0)
                {
                    _logService.LogMessage($"UIタイマーTick: {tempResults.Count}件の結果を追加");
                    foreach (var result in tempResults)
                    {
                        AddSearchResult(result);
                    }
                }
            };

            timer.Start();
            _logService.LogMessage("UIタイマーを開始しました。");
        }

        /// <summary>
        /// 検索キャンセルボタンクリック時の処理
        /// </summary>
        private void BtnCancelSearch_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("検索キャンセルボタンがクリックされました");
            if (_isSearching && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
                UpdateStatus("キャンセル処理中...");
            }
        }

        /// <summary>
        /// グリッド行ダブルクリック時の処理
        /// </summary>
        private void GrdResults_DoubleClick(object sender, EventArgs e)
        {
            _logService.LogMessage("GridView ダブルクリックイベント開始");

            if (grdResults.SelectedRows.Count <= 0)
            {
                _logService.LogMessage("選択行がありません");
                return;
            }

            string filePath = grdResults.SelectedRows[0].Cells["colFilePath"].Value?.ToString();
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                _logService.LogMessage($"ファイルが見つかりません: {filePath}");
                MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Shiftキーが押されているかチェック
            bool isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            _logService.LogMessage($"Shiftキー押下状態: {isShiftPressed}");

            if (isShiftPressed)
            {
                // Shift+ダブルクリック：フォルダを開く
                OpenContainingFolder(filePath);
            }
            else
            {
                // 通常のダブルクリック：Excelファイルを開く
                string sheetName = grdResults.SelectedRows[0].Cells["colSheetName"].Value?.ToString();
                string cellPosition = grdResults.SelectedRows[0].Cells["colCellPosition"].Value?.ToString();
                _logService.LogMessage($"Excel操作を開始します: ファイル={filePath}, シート={sheetName}, セル={cellPosition}");

                // 図形内の場合はセル位置指定なしで開く
                if (cellPosition == "図形内" || cellPosition == "図形内 (GF)")
                {
                    OpenExcel(filePath, sheetName, null);
                }
                else
                {
                    OpenExcel(filePath, sheetName, cellPosition);
                }
            }
        }

        /// <summary>
        /// ファイルの含まれるフォルダを開く
        /// </summary>
        private void OpenContainingFolder(string filePath)
        {
            try
            {
                string folderPath = Path.GetDirectoryName(filePath);
                _logService.LogMessage($"フォルダを開きます: {folderPath}");
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
                UpdateStatus($"{folderPath} フォルダを開きました");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"フォルダを開く際のエラー: {ex.Message}");
                MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Excelファイルを開く
        /// </summary>
        private void OpenExcel(string filePath, string sheetName, string cellPosition)
        {
            bool success = _excelInteropService.OpenExcelFile(filePath, sheetName, cellPosition);

            if (success)
            {
                if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(cellPosition))
                {
                    UpdateStatus($"{Path.GetFileName(filePath)} を開きました。シート '{sheetName}' のセル {cellPosition} を確認してください。");
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                     UpdateStatus($"{Path.GetFileName(filePath)} を開きました。シート '{sheetName}' を確認してください。");
                }
                else
                {
                    UpdateStatus($"{Path.GetFileName(filePath)} を開きました。");
                }
            }
            else
            {
                MessageBox.Show("Excelファイルを開けませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// グリッドでのキー押下時の処理
        /// </summary>
        private void GrdResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                // CTRL+A で全行選択
                _logService.LogMessage("Ctrl+A キーが押されました: 全行選択");
                grdResults.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                // CTRL+C でコピー
                _logService.LogMessage("Ctrl+C キーが押されました: 選択行をコピー");
                CopySelectedRowsToClipboard();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// 選択行をクリップボードにコピー
        /// </summary>
        private void CopySelectedRowsToClipboard()
        {
            _logService.LogMessage("選択行をクリップボードにコピー開始");
            
            if (ExcelUtils.CopySelectedRowsToClipboard(grdResults, message => 
            {
                _logService.LogMessage(message);
                UpdateStatus(message);
            }))
            {
                // 成功時の処理は既にコールバック内で行われている
            }
            else
            {
                if (grdResults.SelectedRows.Count == 0)
                {
                    _logService.LogMessage("コピー対象の行が選択されていません");
                }
                else
                {
                    MessageBox.Show("クリップボードへのコピーに失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// フォーム終了時の処理
        /// </summary>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _logService.LogMessage("アプリケーションを終了します");
            
            if (_isSearching && _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _logService.LogMessage("実行中の検索をキャンセルします");
                _cancellationTokenSource.Cancel();
            }
            
            SaveCurrentSettings();
        }

        /// <summary>
        /// UIの検索状態を設定
        /// </summary>
        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;
            
            // UI要素の有効/無効を切り替え
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
                UpdateStatus("検索中...");
                _logService.LogMessage("検索を開始しました");
            }
            else
            {
                _logService.LogMessage("検索状態を終了しました");
            }
        }

        /// <summary>
        /// 検索条件の入力検証
        /// </summary>
        private bool ValidateSearchInputs()
        {
            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                _logService.LogMessage("フォルダパスが入力されていません");
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(cmbFolderPath.Text))
            {
                _logService.LogMessage($"指定されたフォルダが存在しません: {cmbFolderPath.Text}");
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                _logService.LogMessage("検索キーワードが入力されていません");
                MessageBox.Show("検索キーワードを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 検索結果をグリッドに表示
        /// </summary>
        private void DisplaySearchResults(List<SearchResult> results)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    _logService.LogMessage($"DisplaySearchResults: UIスレッドへの呼び出しが必要です。結果件数={results.Count}");
                    this.Invoke(new Action(() => DisplaySearchResults(results)));
                    return;
                }

                _logService.LogMessage($"DisplaySearchResults: グリッドをクリアして結果を表示します。件数={results.Count}");

                // 結果グリッドをクリア
                grdResults.Rows.Clear();

                // 結果をグリッドに表示
                foreach (var result in results)
                {
                    string fileName = Path.GetFileName(result.FilePath);
                    grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);
                }

                // UIの更新を強制
                grdResults.Refresh();

                _logService.LogMessage($"グリッドに {results.Count} 件の結果を表示しました");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"DisplaySearchResults エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 検索結果を1件追加（リアルタイム表示用）
        /// </summary>
        private void AddSearchResult(SearchResult result)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => AddSearchResult(result)));
                    return;
                }

                string fileName = Path.GetFileName(result.FilePath);
        
                // デバッグ情報を追加
                _logService.LogMessage($"結果追加: {fileName}, シート={result.SheetName}, セル={result.CellPosition}");
        
                int rowIndex = grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);

                // 最新の行にスクロール
                if (grdResults.Rows.Count > 0)
                {
                    grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
                }
        
                // UIの更新を強制
                grdResults.Refresh();
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"AddSearchResult エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// ステータス表示を更新
        /// </summary>
        private void UpdateStatus(string message)
        {
            _logService.UpdateStatus(message);
        }
    }
}