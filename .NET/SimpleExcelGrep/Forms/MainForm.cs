using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
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
            _logService = new LogService(lblStatus, false);
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
            btnLoadTsv.Click += BtnLoadTsv_Click; // TSV読み込みボタンのイベントハンドラ
            grdResults.DoubleClick += GrdResults_DoubleClick;
            grdResults.KeyDown += GrdResults_KeyDown;
            cmbKeyword.KeyDown += CmbKeyword_KeyDown;

            // チェックボックスとNumericUpDownのイベント
            chkRealTimeDisplay.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            nudParallelism.ValueChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            chkFirstHitOnly.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            chkSearchShapes.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            txtIgnoreFileSizeMB.TextChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); }; // 無視ファイルサイズのイベント

            // ===== 機能追加 =====
            chkCellMode.CheckedChanged += ChkCellMode_CheckedChanged;
            txtCellAddress.TextChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
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
                txtIgnoreFileSizeMB.Text = settings.IgnoreFileSizeMB.ToString(CultureInfo.InvariantCulture);
                
                // ===== 機能追加 =====
                chkCellMode.Checked = settings.CellModeEnabled;
                txtCellAddress.Text = settings.CellAddress;
                // UIの状態を初期化
                txtCellAddress.Enabled = chkCellMode.Checked;
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

            double ignoreFileSize = 0;
            if (double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double size))
            {
                ignoreFileSize = size;
            }

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
                SearchShapes = chkSearchShapes.Checked,
                IgnoreFileSizeMB = ignoreFileSize,

                // ===== 機能追加 =====
                CellModeEnabled = chkCellMode.Checked,
                CellAddress = txtCellAddress.Text
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
        /// 検索開始ボタンクリック時の処理 (モード分岐)
        /// </summary>
        private async void BtnStartSearch_Click(object sender, EventArgs e)
        {
            // ① フォルダパスの共通検証
            if (!ValidateFolderPath()) return;

            // 履歴の保存と設定の保存
            _settingsService.AddToComboBoxHistory(cmbFolderPath, cmbFolderPath.Text);
            _settingsService.AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);
            _settingsService.AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);
            SaveCurrentSettings();

            // ② セル検索モードかGREP検索モードかを判定
            if (chkCellMode.Checked && !string.IsNullOrWhiteSpace(txtCellAddress.Text))
            {
                // セル検索モード
                await RunCellSearchModeAsync();
            }
            else
            {
                // GREP検索モード
                await RunGrepSearchModeAsync();
            }
        }

        /// <summary>
        /// GREP検索モードを実行
        /// </summary>
        private async Task RunGrepSearchModeAsync()
        {
            _logService.LogMessage("GREP検索モードを開始します");

            // GREP検索用の入力検証
            if (!ValidateGrepInputs()) return;

            SetSearchingState(true);
            _cancellationTokenSource = new CancellationTokenSource();
            _searchResults.Clear();
            grdResults.Rows.Clear();

            try
            {
                List<string> ignoreKeywords = cmbIgnoreKeywords.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(k => k.Trim()).Where(k => !string.IsNullOrEmpty(k)).ToList();
                Regex regex = null;
                if (chkRegex.Checked)
                {
                    try { regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled | RegexOptions.IgnoreCase); }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"正規表現が無効です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetSearchingState(false); return;
                    }
                }

                bool isRealTimeDisplay = chkRealTimeDisplay.Checked;
                var pendingResults = new ConcurrentQueue<SearchResult>();
                _uiTimer = new System.Windows.Forms.Timer { Interval = 100 };
                StartResultUpdateTimer(_uiTimer, pendingResults, isRealTimeDisplay);

                try
                {
                    double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double ignoreFileSize);

                    _searchResults = await _excelSearchService.SearchExcelFilesAsync(
                        cmbFolderPath.Text, cmbKeyword.Text, chkRegex.Checked, regex,
                        ignoreKeywords, isRealTimeDisplay, chkSearchShapes.Checked, chkFirstHitOnly.Checked,
                        (int)nudParallelism.Value, ignoreFileSize, pendingResults, UpdateStatus,
                        _cancellationTokenSource.Token);
                    
                    DisplaySearchResults(_searchResults);
                    UpdateStatus($"検索完了: {_searchResults.Count} 件見つかりました");
                    if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
                }
                catch (OperationCanceledException)
                {
                    UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました");
                    if (!isRealTimeDisplay) DisplaySearchResults(_searchResults);
                    if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus("エラーが発生しました");
                    if (!isRealTimeDisplay) DisplaySearchResults(_searchResults);
                }
                finally
                {
                    if (_uiTimer != null) { _uiTimer.Stop(); _uiTimer.Dispose(); _uiTimer = null; }
                }
            }
            finally
            {
                SetSearchingState(false);
            }
        }

        /// <summary>
        /// セル検索モードを実行
        /// </summary>
        private async Task RunCellSearchModeAsync()
        {
            _logService.LogMessage("セル検索モードを開始します");

            // ⑤ セルアドレスの検証
            string[] cellAddresses = txtCellAddress.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                      .Select(a => a.Trim().ToUpper())
                                                      .ToArray();
            if (cellAddresses.Length == 0 || !cellAddresses.All(IsValidCellAddress))
            {
                _logService.LogMessage($"セルアドレスの形式が不正です: {txtCellAddress.Text}");
                MessageBox.Show("セルアドレスの形式が正しくありません。\nA1形式で、カンマ区切りで入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SetSearchingState(true);
            _cancellationTokenSource = new CancellationTokenSource();
            _searchResults.Clear();
            grdResults.Rows.Clear();

            try
            {
                List<string> ignoreKeywords = cmbIgnoreKeywords.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(k => k.Trim()).Where(k => !string.IsNullOrEmpty(k)).ToList();
                double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double ignoreFileSize);

                bool isRealTimeDisplay = chkRealTimeDisplay.Checked;
                var pendingResults = new ConcurrentQueue<SearchResult>();
                _uiTimer = new System.Windows.Forms.Timer { Interval = 100 };
                StartResultUpdateTimer(_uiTimer, pendingResults, isRealTimeDisplay);
                
                try
                {
                     _searchResults = await _excelSearchService.SearchCellsByAddressAsync(
                        cmbFolderPath.Text, cellAddresses, ignoreKeywords,
                        (int)nudParallelism.Value, ignoreFileSize, pendingResults, UpdateStatus,
                        _cancellationTokenSource.Token);

                    DisplaySearchResults(_searchResults);
                    UpdateStatus($"検索完了: {_searchResults.Count} 件見つかりました");
                    if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
                }
                catch (OperationCanceledException)
                {
                    UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました");
                    if (!isRealTimeDisplay) DisplaySearchResults(_searchResults);
                    if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    UpdateStatus("エラーが発生しました");
                    if (!isRealTimeDisplay) DisplaySearchResults(_searchResults);
                }
                finally
                {
                    if (_uiTimer != null) { _uiTimer.Stop(); _uiTimer.Dispose(); _uiTimer = null; }
                }
            }
            finally
            {
                SetSearchingState(false);
            }
        }
        
        /// <summary>
        /// 結果更新用タイマーを開始
        /// </summary>
        private void StartResultUpdateTimer(System.Windows.Forms.Timer timer, ConcurrentQueue<SearchResult> pendingResults, bool isRealTimeDisplay)
        {
            timer.Tick += (s, e) =>
            {
                const int MaxUpdatesPerTick = 100;
                List<SearchResult> tempResults = new List<SearchResult>();

                int batchSize = 0;
                while (batchSize < MaxUpdatesPerTick && pendingResults.TryDequeue(out SearchResult result))
                {
                    tempResults.Add(result);
                    batchSize++;
                }

                if (isRealTimeDisplay && tempResults.Count > 0)
                {
                    _logService.LogMessage($"UIタイマーTick: {tempResults.Count}件の結果を追加");
                    foreach (var result in tempResults)
                    {
                        _searchResults.Add(result);
                        AddSearchResultToGrid(result);
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
        /// TSV読み込みボタンクリック時の処理
        /// </summary>
        private void BtnLoadTsv_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("TSV読み込みボタンがクリックされました");

            if (_isSearching)
            {
                MessageBox.Show("検索中はTSVファイルを読み込めません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (grdResults.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "現在の検索結果はクリアされます。TSVファイルを読み込みますか？",
                    "確認",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (dialogResult == DialogResult.No)
                {
                    _logService.LogMessage("TSV読み込みはキャンセルされました。");
                    return;
                }
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "TSV files (*.tsv)|*.tsv|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    _logService.LogMessage($"TSVファイルを選択: {filePath}");
                    LoadTsvFile(filePath);
                }
                else
                {
                    _logService.LogMessage("TSVファイル選択はキャンセルされました。");
                }
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

            bool isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            _logService.LogMessage($"Shiftキー押下状態: {isShiftPressed}");

            if (isShiftPressed)
            {
                OpenContainingFolder(filePath);
            }
            else
            {
                string sheetName = grdResults.SelectedRows[0].Cells["colSheetName"].Value?.ToString();
                string cellPosition = grdResults.SelectedRows[0].Cells["colCellPosition"].Value?.ToString();
                _logService.LogMessage($"Excel操作を開始します: ファイル={filePath}, シート={sheetName}, セル={cellPosition}");

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
                _logService.LogMessage("Ctrl+A キーが押されました: 全行選択");
                grdResults.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
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
            if (!ExcelUtils.CopySelectedRowsToClipboard(grdResults, message =>
            {
                _logService.LogMessage(message);
                UpdateStatus(message);
            }))
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

            // 共通UI
            cmbFolderPath.Enabled = !isSearching;
            btnSelectFolder.Enabled = !isSearching;
            btnStartSearch.Enabled = !isSearching;
            btnLoadTsv.Enabled = !isSearching;
            btnCancelSearch.Enabled = isSearching;
            
            // GREP検索用UI
            cmbKeyword.Enabled = !isSearching;
            cmbIgnoreKeywords.Enabled = !isSearching;
            txtIgnoreFileSizeMB.Enabled = !isSearching;
            chkRegex.Enabled = !isSearching;
            chkSearchShapes.Enabled = !isSearching;
            chkFirstHitOnly.Enabled = !isSearching;
            nudParallelism.Enabled = !isSearching;

            // セル検索モード用UI
            chkCellMode.Enabled = !isSearching;
            txtCellAddress.Enabled = !isSearching && chkCellMode.Checked; // 検索中でなく、かつチェックボックスがONの時だけ有効

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
        /// フォルダパスの検証 (共通)
        /// </summary>
        private bool ValidateFolderPath()
        {
            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!Directory.Exists(cmbFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        /// <summary>
        /// GREP検索の入力検証
        /// </summary>
        private bool ValidateGrepInputs()
        {
            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            if (!string.IsNullOrWhiteSpace(txtIgnoreFileSizeMB.Text))
            {
                if (!double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double fileSize) || fileSize < 0)
                {
                    MessageBox.Show("無視ファイルサイズには0以上の数値を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            else
            {
                txtIgnoreFileSizeMB.Text = "0";
            }
            return true;
        }

        /// <summary>
        /// セルアドレスの形式を検証 (A1形式)
        /// </summary>
        private bool IsValidCellAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return false;
            // 簡易的なA1形式の正規表現 (例: A1, BZ10, ABC1048576)
            return Regex.IsMatch(address, @"^[A-Z]{1,3}[1-9][0-9]{0,6}$", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// セルモードチェックボックス変更時の処理
        /// </summary>
        private void ChkCellMode_CheckedChanged(object sender, EventArgs e)
        {
            // ① チェックボックスの状態に応じてテキストボックスの有効/無効を切り替える
            txtCellAddress.Enabled = chkCellMode.Checked;
            if (!_isLoading)
            {
                SaveCurrentSettings();
            }
        }

        /// <summary>
        /// 検索結果をグリッドに表示 (リスト全体)
        /// </summary>
        private void DisplaySearchResults(List<SearchResult> results)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => DisplaySearchResults(results)));
                    return;
                }
                grdResults.Rows.Clear();
                foreach (var result in results)
                {
                    AddSearchResultToGrid(result);
                }
                grdResults.Refresh();
                _logService.LogMessage($"グリッドに {results.Count} 件の結果を表示しました");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"DisplaySearchResults エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 検索結果を1件グリッドに追加
        /// </summary>
        private void AddSearchResultToGrid(SearchResult result)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(() => AddSearchResultToGrid(result)));
                    return;
                }

                string fileName = "";
                try { fileName = Path.GetFileName(result.FilePath); }
                catch (ArgumentException ex)
                {
                    _logService.LogMessage($"ファイルパスからファイル名取得エラー: {result.FilePath}, {ex.Message}");
                    fileName = result.FilePath;
                }

                int rowIndex = grdResults.Rows.Add(result.FilePath, fileName, result.SheetName, result.CellPosition, result.CellValue);

                if (_isSearching && chkRealTimeDisplay.Checked && grdResults.Rows.Count > 0)
                {
                    grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
                }
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"AddSearchResultToGrid エラー: {ex.Message}");
            }
        }


        /// <summary>
        /// ステータス表示を更新
        /// </summary>
        private void UpdateStatus(string message)
        {
            _logService.UpdateStatus(message);
        }

        /// <summary>
        /// 検索結果をTSVファイルに書き出す
        /// </summary>
        private void WriteResultsToTsv(List<SearchResult> results)
        {
            if (results == null || !results.Any())
            {
                _logService.LogMessage("TSV書き出し: 書き出す結果がありません。");
                return;
            }

            try
            {
                string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string fileName = $"{DateTime.Now:yyyyMMdd_HHmmss}.tsv";
                string filePath = Path.Combine(exePath, fileName);

                _logService.LogMessage($"TSVファイル書き出し開始: {filePath}");

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("ファイルパス\tファイル名\tシート名\tセル位置\tセルの値");

                foreach (var result in results)
                {
                    string tsvRow = string.Join("\t",
                        EscapeTsvField(result.FilePath),
                        EscapeTsvField(Path.GetFileName(result.FilePath)),
                        EscapeTsvField(result.SheetName),
                        EscapeTsvField(result.CellPosition),
                        EscapeTsvField(result.CellValue)
                    );
                    sb.AppendLine(tsvRow);
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                _logService.LogMessage($"TSVファイル書き出し完了: {results.Count}件");
                UpdateStatus($"TSVファイルに結果を書き出しました: {fileName}");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル書き出しエラー: {ex.Message}");
                MessageBox.Show($"TSVファイルへの書き出し中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// TSVフィールドのエスケープ処理
        /// </summary>
        private string EscapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            if (field.Contains('\t') || field.Contains('\n') || field.Contains('\r') || field.Contains('"'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }


        /// <summary>
        /// TSVファイルを読み込み、DataGridViewに表示する
        /// </summary>
        private void LoadTsvFile(string filePath)
        {
            try
            {
                _logService.LogMessage($"TSVファイル読み込み開始: {filePath}");
                grdResults.Rows.Clear();
                _searchResults.Clear();

                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);

                if (lines.Length <= 1)
                {
                    _logService.LogMessage("TSVファイルにデータがありません。");
                    MessageBox.Show("TSVファイルに読み込むデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                for (int i = 1; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string[] fields = ParseTsvLine(line);

                    if (fields.Length >= 5)
                    {
                        SearchResult result = new SearchResult
                        {
                            FilePath = UnescapeTsvField(fields[0]),
                            SheetName = UnescapeTsvField(fields[2]),
                            CellPosition = UnescapeTsvField(fields[3]),
                            CellValue = UnescapeTsvField(fields[4])
                        };
                        _searchResults.Add(result);
                        AddSearchResultToGrid(result);
                    }
                    else
                    {
                        _logService.LogMessage($"TSV行の形式が不正です (列数不足): {line}");
                    }
                }
                grdResults.Refresh();
                UpdateStatus($"TSVファイルから {_searchResults.Count} 件の結果を読み込みました。");
                _logService.LogMessage($"TSVファイル読み込み完了: {_searchResults.Count}件");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル読み込みエラー: {ex.Message}");
                MessageBox.Show($"TSVファイルの読み込み中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("TSVファイルの読み込みに失敗しました。");
            }
        }

        /// <summary>
        /// TSVの1行をパースする
        /// </summary>
        private string[] ParseTsvLine(string line)
        {
            List<string> fields = new List<string>();
            StringBuilder currentField = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        currentField.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == '\t' && !inQuotes)
                {
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            fields.Add(currentField.ToString());
            return fields.ToArray();
        }

        /// <summary>
        /// TSVフィールドのアンエスケープ処理
        /// </summary>
        private string UnescapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";

            if (field.StartsWith("\"") && field.EndsWith("\""))
            {
                string unescaped = field.Substring(1, field.Length - 2);
                return unescaped.Replace("\"\"", "\"");
            }
            return field;
        }
    }
}