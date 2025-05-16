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
using System.Diagnostics; // For Process.GetCurrentProcess()

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
        private bool _isLoadingSettings = false; // 設定ロード中フラグ
        private List<SearchResult> _searchResults = new List<SearchResult>();
        private System.Windows.Forms.Timer _uiUpdateTimer; // UI更新用タイマー
        private ConcurrentQueue<SearchResult> _pendingResultsQueue; // 検索結果を一時的に保持するキュー


        public MainForm()
        {
            InitializeComponent();

            // サービスの初期化
            // LogService の isLoggingEnabled を true にしてログを有効化
            _logService = new LogService(lblStatus, true);
            _settingsService = new SettingsService(_logService);
            _excelSearchService = new ExcelSearchService(_logService);
            _excelInteropService = new ExcelInteropService(_logService);

            // イベントハンドラの登録
            RegisterEventHandlers();

            // UIの初期設定
            InitializeUI();
            UpdateWindowTitleWithVersion();
        }

        /// <summary>
        /// アプリケーションのバージョン情報をウィンドウタイトルに表示
        /// </summary>
        private void UpdateWindowTitleWithVersion()
        {
            try
            {
                Version version = Assembly.GetExecutingAssembly().GetName().Version;
                this.Text = $"Simple Excel Grep v{version.Major}.{version.Minor}.{version.Build}";
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"バージョン情報の取得に失敗: {ex.Message}");
                this.Text = "Simple Excel Grep";
            }
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
            btnLoadTsv.Click += BtnLoadTsv_Click;
            grdResults.DoubleClick += GrdResults_DoubleClick;
            grdResults.KeyDown += GrdResults_KeyDown;
            cmbKeyword.KeyDown += CmbKeyword_KeyDown;

            // UI要素の変更イベントで設定を保存 (ロード時以外)
            chkRealTimeDisplay.CheckedChanged += (s, e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };
            nudParallelism.ValueChanged += (s, e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };
            chkFirstHitOnly.CheckedChanged += (s, e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };
            chkSearchShapes.CheckedChanged += (s, e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };
            txtIgnoreFileSizeMB.TextChanged += (s, e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };
            chkRegex.CheckedChanged += (s,e) => { if (!_isLoadingSettings) SaveCurrentSettings(); };

            // コンボボックスのテキスト変更完了時にも設定を保存するようにする（履歴のため）
            cmbFolderPath.Leave += (s, e) => { if (!_isLoadingSettings && !string.IsNullOrWhiteSpace(cmbFolderPath.Text)) SaveCurrentSettingsWithHistoryUpdate(cmbFolderPath, cmbFolderPath.Text); };
            cmbKeyword.Leave += (s, e) => { if (!_isLoadingSettings && !string.IsNullOrWhiteSpace(cmbKeyword.Text)) SaveCurrentSettingsWithHistoryUpdate(cmbKeyword, cmbKeyword.Text); };
            cmbIgnoreKeywords.Leave += (s, e) => { if (!_isLoadingSettings && !string.IsNullOrWhiteSpace(cmbIgnoreKeywords.Text)) SaveCurrentSettingsWithHistoryUpdate(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);};
        }

        private void SaveCurrentSettingsWithHistoryUpdate(ComboBox comboBox, string newText)
        {
            // 実際に検索実行時やアプリ終了時にも保存されるので、ここでは履歴更新を主目的とする
             _settingsService.AddToComboBoxHistory(comboBox, newText); // 履歴に追加・更新
            SaveCurrentSettings(); // 通常の設定保存も行う
        }


        /// <summary>
        /// UIの初期設定
        /// </summary>
        private void InitializeUI()
        {
            grdResults.MultiSelect = true;
            grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grdResults.RowHeadersVisible = false; // 行ヘッダーを非表示に
            grdResults.AllowUserToResizeRows = false; // 行の高さ変更を禁止

            // カラム幅の調整 (手動で調整、またはFillWeightで調整)
            colFilePath.Visible = false; // ファイルパス全体は非表示にすることが多い（ファイル名で十分な場合）
            colFileName.FillWeight = 30;      // ファイル名
            colSheetName.FillWeight = 20;     // シート名
            colCellPosition.FillWeight = 10;  // セル位置
            colCellValue.FillWeight = 40;     // セルの値 (これが一番重要なので広めに)

            ContextMenuStrip contextMenu = new ContextMenuStrip();
            ToolStripMenuItem copyMenuItem = new ToolStripMenuItem("コピー (Ctrl+C)");
            copyMenuItem.Click += (s, e) => CopySelectedRowsToClipboard();
            contextMenu.Items.Add(copyMenuItem);

            ToolStripMenuItem selectAllMenuItem = new ToolStripMenuItem("すべて選択 (Ctrl+A)");
            selectAllMenuItem.Click += (s, e) => grdResults.SelectAll();
            contextMenu.Items.Add(selectAllMenuItem);
            contextMenu.Items.Add(new ToolStripSeparator());
            ToolStripMenuItem openFileMenuItem = new ToolStripMenuItem("ファイルを開く (ダブルクリック)");
            openFileMenuItem.Click += (s,e) => OpenSelectedFileFromGrid();
            contextMenu.Items.Add(openFileMenuItem);
            ToolStripMenuItem openFolderMenuItem = new ToolStripMenuItem("フォルダを開く (Shift+ダブルクリック)");
            openFolderMenuItem.Click += (s,e) => OpenSelectedFileContainingFolderFromGrid();
            contextMenu.Items.Add(openFolderMenuItem);


            grdResults.ContextMenuStrip = contextMenu;

            // UI更新用タイマーの初期化
            _pendingResultsQueue = new ConcurrentQueue<SearchResult>();
            _uiUpdateTimer = new System.Windows.Forms.Timer();
            _uiUpdateTimer.Interval = 200; // 200ミリ秒ごとにUI更新（適宜調整）
            _uiUpdateTimer.Tick += UiUpdateTimer_Tick;
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
            _isLoadingSettings = true;
            _logService.LogMessage("設定の読み込み開始 (LoadSettings)");

            try
            {
                Settings settings = _settingsService.LoadSettings();

                cmbFolderPath.Items.Clear();
                foreach (var path in settings.FolderPathHistory.Where(p => !string.IsNullOrWhiteSpace(p))) cmbFolderPath.Items.Add(path);
                cmbFolderPath.Text = settings.FolderPath;

                cmbKeyword.Items.Clear();
                foreach (var keyword in settings.SearchKeywordHistory.Where(k => !string.IsNullOrWhiteSpace(k))) cmbKeyword.Items.Add(keyword);
                cmbKeyword.Text = settings.SearchKeyword;

                cmbIgnoreKeywords.Items.Clear();
                foreach (var keyword in settings.IgnoreKeywordsHistory.Where(k => !string.IsNullOrWhiteSpace(k))) cmbIgnoreKeywords.Items.Add(keyword);
                cmbIgnoreKeywords.Text = settings.IgnoreKeywords;

                chkRegex.Checked = settings.UseRegex;
                chkRealTimeDisplay.Checked = settings.RealTimeDisplay;
                nudParallelism.Value = Math.Min(Math.Max(settings.MaxParallelism, nudParallelism.Minimum), nudParallelism.Maximum);
                chkFirstHitOnly.Checked = settings.FirstHitOnly;
                chkSearchShapes.Checked = settings.SearchShapes;
                txtIgnoreFileSizeMB.Text = settings.IgnoreFileSizeMB.ToString(CultureInfo.InvariantCulture);

                _logService.LogMessage("設定の読み込み完了 (LoadSettings)");
            }
            catch (Exception ex)
            {
                 _logService.LogMessage($"設定の読み込み中にエラー: {ex.Message}");
            }
            finally
            {
                _isLoadingSettings = false;
            }
        }

        /// <summary>
        /// 現在のUI設定を保存
        /// </summary>
        private void SaveCurrentSettings()
        {
            if (_isLoadingSettings) return; // 設定ロード中は保存しない
            _logService.LogMessage("設定の保存開始 (SaveCurrentSettings)");


            double ignoreFileSize = 0;
            if (double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double size))
            {
                ignoreFileSize = size;
            }
            else if (!string.IsNullOrWhiteSpace(txtIgnoreFileSizeMB.Text)) // パース失敗かつ空でもない場合
            {
                 _logService.LogMessage($"無視ファイルサイズの入力が無効です: '{txtIgnoreFileSizeMB.Text}'。保存時は0として扱います。");
                 // Optionally, notify the user or revert to a valid value
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
                IgnoreFileSizeMB = ignoreFileSize
            };

            if (!_settingsService.SaveSettings(settings))
            {
                MessageBox.Show("設定の保存に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                _logService.LogMessage("設定の保存完了 (SaveCurrentSettings)");
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
                dialog.ShowNewFolderButton = false; // 通常、既存フォルダを選択

                if (!string.IsNullOrEmpty(cmbFolderPath.Text) && Directory.Exists(cmbFolderPath.Text))
                {
                    dialog.SelectedPath = cmbFolderPath.Text;
                }
                else // 前回値が無効ならデフォルトパス（例：ドキュメント）を設定しても良い
                {
                    dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }


                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _logService.LogMessage($"選択されたフォルダ: {dialog.SelectedPath}");
                    // cmbFolderPath.Text = dialog.SelectedPath; // これで履歴更新と設定保存がトリガーされるはず
                    _settingsService.AddToComboBoxHistory(cmbFolderPath, dialog.SelectedPath); //明示的に履歴更新
                    SaveCurrentSettings(); // 設定保存
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
                _logService.LogMessage("検索キーワード入力中にEnterキー押下、検索を開始します。");
                BtnStartSearch_Click(this, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// 検索開始ボタンクリック時の処理
        /// </summary>
        private async void BtnStartSearch_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("検索開始ボタンがクリックされました。");

            if (!ValidateSearchInputs()) return;

            _settingsService.AddToComboBoxHistory(cmbFolderPath, cmbFolderPath.Text);
            _settingsService.AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);
            _settingsService.AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);
            SaveCurrentSettings();

            _logService.InitializePerformanceLog();

            SetSearchingState(true);
            _cancellationTokenSource = new CancellationTokenSource();
            _searchResults.Clear();
            _pendingResultsQueue = new ConcurrentQueue<SearchResult>(); // キューをクリア
            grdResults.Rows.Clear();
            if (chkRealTimeDisplay.Checked) _uiUpdateTimer.Start(); // リアルタイム表示の場合タイマー開始


            try
            {
                List<string> ignoreKeywords = cmbIgnoreKeywords.Text
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(k => k.Trim().ToLowerInvariant()) // 小文字で統一して比較容易に
                    .Where(k => !string.IsNullOrEmpty(k))
                    .ToList();
                _logService.LogMessage($"無視キーワードリスト (小文字化): [{string.Join(", ", ignoreKeywords)}]");

                Regex regex = null;
                if (chkRegex.Checked)
                {
                    try
                    {
                        regex = new Regex(cmbKeyword.Text, RegexOptions.Compiled | RegexOptions.IgnoreCase);
                        _logService.LogMessage($"正規表現コンパイル成功: {cmbKeyword.Text}");
                    }
                    catch (ArgumentException ex) // Regexコンパイルエラー
                    {
                        _logService.LogMessage($"正規表現が無効です: {ex.Message}");
                        MessageBox.Show($"正規表現が無効です: {ex.Message}", "正規表現エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        SetSearchingState(false);
                        return;
                    }
                }

                double ignoreFileSize = 0;
                if (!double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out ignoreFileSize) || ignoreFileSize < 0)
                {
                    ignoreFileSize = 0; // 不正な値は0として扱う（警告はValidateSearchInputsで出す）
                }

                _logService.LogMessage($"現在のメモリ使用量 (検索タスク開始直前): {GetCurrentMemoryUsageMB():F2} MB");

                // 検索タスクの実行
                _searchResults = await _excelSearchService.SearchExcelFilesAsync(
                    cmbFolderPath.Text,
                    cmbKeyword.Text,
                    chkRegex.Checked,
                    regex,
                    ignoreKeywords,
                    chkRealTimeDisplay.Checked,
                    chkSearchShapes.Checked,
                    chkFirstHitOnly.Checked,
                    (int)nudParallelism.Value,
                    ignoreFileSize,
                    _pendingResultsQueue, // UI更新用キュー
                    UpdateStatus, // ステータス更新コールバック
                    _cancellationTokenSource.Token
                );

                // 検索タスク完了後、キューに残っているものを処理 (リアルタイム表示OFFの場合や、タイミングによっては残る可能性)
                ProcessPendingResultsQueue(); // UIタイマーが止まっていても確実に処理

                if (!_cancellationTokenSource.IsCancellationRequested)
                {
                     UpdateStatus($"検索完了: {_searchResults.Count} 件見つかりました。");
                     _logService.LogMessage($"検索正常完了。総ヒット数: {_searchResults.Count}");
                }
                else
                {
                    UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました。");
                    _logService.LogMessage($"検索中止完了。その時点でのヒット数: {_searchResults.Count}");
                }


                // 最終結果をグリッドに表示 (リアルタイム表示OFFの場合、ここでまとめて表示)
                if (!chkRealTimeDisplay.Checked)
                {
                    _logService.LogMessage("リアルタイム表示OFFのため、結果をまとめてグリッドに表示します。");
                    DisplaySearchResults(_searchResults);
                }


                if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
                _logService.LogMessage($"現在のメモリ使用量 (検索タスク完了後): {GetCurrentMemoryUsageMB():F2} MB");

            }
            catch (OperationCanceledException) // SearchExcelFilesAsync からのキャンセル
            {
                UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました。");
                _logService.LogMessage("検索処理がユーザーによってキャンセルされました (BtnStartSearch_Click catch)。");
                ProcessPendingResultsQueue(); // 中断時もキューの残りを処理
                if (!chkRealTimeDisplay.Checked) DisplaySearchResults(_searchResults);
                if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"検索中に予期せぬエラーが発生しました (BtnStartSearch_Click catch): {ex.ToString()}");
                MessageBox.Show($"検索中に重大なエラーが発生しました: {ex.Message}\n詳細はログファイルを確認してください。", "検索エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("エラーが発生しました。");
                ProcessPendingResultsQueue(); // エラー時もキューの残りを処理
                if (!chkRealTimeDisplay.Checked) DisplaySearchResults(_searchResults);

            }
            finally
            {
                if (_uiUpdateTimer.Enabled) _uiUpdateTimer.Stop();
                SetSearchingState(false);
                _logService.LogMessage("検索処理の finally ブロック完了。");
            }
        }


        /// <summary>
        /// 結果更新用タイマーのTickイベント
        /// </summary>
        private void UiUpdateTimer_Tick(object sender, EventArgs e)
        {
            ProcessPendingResultsQueue();
        }

        /// <summary>
        /// 保留中の検索結果キューを処理してUIに反映
        /// </summary>
        private void ProcessPendingResultsQueue()
        {
            if (_pendingResultsQueue == null) return;

            const int maxItemsToProcessPerTick = 100; // 一度に処理する最大件数 (UIの応答性を保つため)
            int processedCount = 0;
            List<SearchResult> newResultsForGrid = new List<SearchResult>();

            while (processedCount < maxItemsToProcessPerTick && _pendingResultsQueue.TryDequeue(out SearchResult result))
            {
                // _searchResults は検索タスク完了後にawaitの結果でまとめて更新されるか、
                // リアルタイム表示の場合はここで追加する。今回はawaitの結果を使うので、ここではUI用リストのみ。
                // ただし、最終的な結果リスト(_searchResults)にも追加が必要ならここで行う。
                // 今回の設計では ExcelSearchService の結果が直接 _searchResults に入るので、
                // このキューは純粋にUI更新のバッファとなる。
                // → 設計変更：_searchResultsにもここで追加する。ExcelSearchServiceはキューに入れるだけ。
                _searchResults.Add(result);
                newResultsForGrid.Add(result);
                processedCount++;
            }

            if (newResultsForGrid.Any())
            {
                AddSearchResultsToGrid(newResultsForGrid); // 複数件まとめて追加
                UpdateStatus($"検索中... {_searchResults.Count} 件発見 (処理中ファイル {_searchResults.GroupBy(r => r.FilePath).Count()})");
            }
        }


        /// <summary>
        /// 検索キャンセルボタンクリック時の処理
        /// </summary>
        private void BtnCancelSearch_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("検索キャンセルボタンがクリックされました。");
            if (_isSearching && _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _logService.LogMessage("キャンセル処理を開始します...");
                _cancellationTokenSource.Cancel();
                UpdateStatus("キャンセル処理中...");
                // SetSearchingState(false) は検索タスクのfinallyで行う
            }
            else
            {
                _logService.LogMessage("検索中でないか、既にキャンセル処理が開始されています。");
            }
        }

        /// <summary>
        /// TSV読み込みボタンクリック時の処理
        /// </summary>
        private void BtnLoadTsv_Click(object sender, EventArgs e)
        {
            _logService.LogMessage("TSV読み込みボタンがクリックされました。");

            if (_isSearching)
            {
                MessageBox.Show("検索中はTSVファイルを読み込めません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (grdResults.Rows.Count > 0)
            {
                DialogResult dialogResult = MessageBox.Show(
                    "現在の検索結果はクリアされます。TSVファイルを読み込みますか？",
                    "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.No)
                {
                    _logService.LogMessage("TSV読み込みはユーザーによってキャンセルされました。");
                    return;
                }
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "TSV files (*.tsv)|*.tsv|All files (*.*)|*.*";
                openFileDialog.Title = "読み込むTSVファイルを選択してください";
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    LoadTsvFile(openFileDialog.FileName);
                }
                else
                {
                    _logService.LogMessage("TSVファイル選択はキャンセルされました。");
                }
            }
        }


        /// <summary>
        /// グリッド行ダブルクリック時の処理 (ファイルを開く)
        /// </summary>
        private void GrdResults_DoubleClick(object sender, EventArgs e)
        {
            OpenSelectedFileFromGrid((Control.ModifierKeys & Keys.Shift) == Keys.Shift);
        }

        private void OpenSelectedFileFromGrid(bool openFolder = false)
        {
            if (grdResults.SelectedRows.Count <= 0)
            {
                _logService.LogMessage("ファイルを開こうとしましたが、行が選択されていません。");
                return;
            }

            // 最初の選択行の情報を取得
            DataGridViewRow selectedRow = grdResults.SelectedRows[0];
            string filePath = selectedRow.Cells["colFilePath"].Value?.ToString();

            if (string.IsNullOrEmpty(filePath))
            {
                _logService.LogMessage($"ファイルパスが無効です (null or empty)。");
                MessageBox.Show("ファイルパスが無効です。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!File.Exists(filePath))
            {
                 _logService.LogMessage($"ファイルが見つかりません: {filePath}");
                 MessageBox.Show($"指定されたファイルが見つかりません:\n{filePath}", "ファイル未検出", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                 return;
            }


            if (openFolder)
            {
                OpenContainingFolder(filePath);
            }
            else
            {
                string sheetName = selectedRow.Cells["colSheetName"].Value?.ToString();
                string cellPosition = selectedRow.Cells["colCellPosition"].Value?.ToString();
                OpenExcel(filePath, sheetName, cellPosition);
            }
        }
         private void OpenSelectedFileContainingFolderFromGrid()
        {
            OpenSelectedFileFromGrid(true);
        }


        /// <summary>
        /// ファイルの含まれるフォルダを開く
        /// </summary>
        private void OpenContainingFolder(string filePath)
        {
            try
            {
                string folderPath = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                {
                    _logService.LogMessage($"フォルダパスが無効または存在しません: {folderPath}");
                    MessageBox.Show("フォルダパスが無効か、フォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                _logService.LogMessage($"フォルダを開きます: {folderPath}");
                Process.Start("explorer.exe", $"/select,\"{filePath}\""); // ファイルを選択状態で開く
                UpdateStatus($"フォルダを開きました: {folderPath}");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"フォルダを開く際にエラー: {ex.Message}");
                MessageBox.Show($"フォルダを開けませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Excelファイルを開く (Interop経由)
        /// </summary>
        private void OpenExcel(string filePath, string sheetName, string cellPosition)
        {
            _logService.LogMessage($"Excelファイルを開く準備: ファイル='{filePath}', シート='{sheetName}', セル='{cellPosition}'");

            // 図形内の場合はセル位置指定なしで開く
            if (cellPosition == "図形内" || cellPosition == "図形内 (GF)")
            {
                _logService.LogMessage("図形内ヒットのため、セル位置指定なしで開きます。");
                cellPosition = null; // または ""
            }

            bool success = _excelInteropService.OpenExcelFile(filePath, sheetName, cellPosition);

            if (success)
            {
                string msg = $"{Path.GetFileName(filePath)} を開きました。";
                if (!string.IsNullOrEmpty(sheetName)) msg += $" シート '{sheetName}'";
                if (!string.IsNullOrEmpty(cellPosition)) msg += $" セル {cellPosition}";
                msg += " を確認してください。";
                UpdateStatus(msg);
                _logService.LogMessage("Excelファイルのオープンに成功しました。");
            }
            else
            {
                _logService.LogMessage("Excelファイルのオープンに失敗しました。");
                // ExcelInteropService内でエラーメッセージはログされているはず
                // MessageBox.Show("Excelファイルを開けませんでした。詳細はログを確認してください。", "Excelオープンエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // ExcelInteropService内で既にProcess.Startで開く試みもしているので、ここでは追加のMessageBoxは不要かもしれない。
            }
        }

        /// <summary>
        /// グリッドでのキー押下時の処理 (Ctrl+A, Ctrl+C)
        /// </summary>
        private void GrdResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                _logService.LogMessage("Ctrl+A キー押下: 全行選択");
                grdResults.SelectAll();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                _logService.LogMessage("Ctrl+C キー押下: 選択行をコピー");
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
            _logService.LogMessage("選択行のクリップボードへのコピー処理を開始します。");
            if (grdResults.SelectedRows.Count == 0)
            {
                _logService.LogMessage("コピーする行が選択されていません。");
                UpdateStatus("コピーする行が選択されていません。");
                return;
            }

            if (ExcelUtils.CopySelectedRowsToClipboard(grdResults, message =>
            {
                _logService.LogMessage(message); // ExcelUtilsからのログメッセージ
                UpdateStatus(message);           // UIにも表示
            }))
            {
                _logService.LogMessage("クリップボードへのコピーが成功しました。");
            }
            else
            {
                _logService.LogMessage("クリップボードへのコピーに失敗しました。");
                // ExcelUtils内で具体的なエラーメッセージはログされているはず
                // UpdateStatus("コピーに失敗しました。"); // ExcelUtils側でメッセージ更新している
            }
        }

        /// <summary>
        /// フォーム終了時の処理
        /// </summary>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _logService.LogMessage("アプリケーション終了処理開始 (FormClosing)。");

            if (_isSearching && _cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
            {
                _logService.LogMessage("実行中の検索があるため、キャンセルを試みます...");
                _cancellationTokenSource.Cancel();
                // 非同期処理の完了を待つか、ユーザーに確認するか検討
                // ここではキャンセル要求だけして、設定保存に進む
                // e.Cancel = true; // 終了を一旦キャンセルして、検索の完了を待つ場合
            }

            // UIタイマーが動いていれば停止
            if (_uiUpdateTimer != null && _uiUpdateTimer.Enabled)
            {
                _uiUpdateTimer.Stop();
                _uiUpdateTimer.Dispose();
                _logService.LogMessage("UI更新タイマーを停止・破棄しました。");
            }


            SaveCurrentSettings(); // 最後に設定を保存
            _logService.LogMessage("アプリケーションを終了します。");
        }

        /// <summary>
        /// UIの検索状態を設定 (ボタンの有効/無効など)
        /// </summary>
        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;
            _logService.LogMessage($"検索状態を '{isSearching}' に設定します。");


            // UI要素の有効/無効をスレッドセーフに切り替え
            Action updateUiAction = () =>
            {
                cmbFolderPath.Enabled = !isSearching;
                cmbKeyword.Enabled = !isSearching;
                cmbIgnoreKeywords.Enabled = !isSearching;
                txtIgnoreFileSizeMB.Enabled = !isSearching;
                chkRegex.Enabled = !isSearching;
                chkSearchShapes.Enabled = !isSearching;
                chkFirstHitOnly.Enabled = !isSearching;
                chkRealTimeDisplay.Enabled = !isSearching;
                nudParallelism.Enabled = !isSearching;

                btnSelectFolder.Enabled = !isSearching;
                btnStartSearch.Enabled = !isSearching;
                btnLoadTsv.Enabled = !isSearching;
                btnCancelSearch.Enabled = isSearching;

                if (isSearching)
                {
                    UpdateStatus("検索中...");
                    Cursor = Cursors.WaitCursor;
                }
                else
                {
                    // UpdateStatus("準備完了"); // 検索完了メッセージは別途設定
                    Cursor = Cursors.Default;
                }
            };

            if (this.InvokeRequired)
            {
                this.Invoke(updateUiAction);
            }
            else
            {
                updateUiAction();
            }
        }

        /// <summary>
        /// 検索条件の入力検証
        /// </summary>
        private bool ValidateSearchInputs()
        {
            _logService.LogMessage("検索条件の入力検証を開始します。");
            if (string.IsNullOrWhiteSpace(cmbFolderPath.Text))
            {
                MessageBox.Show("検索対象のフォルダパスを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbFolderPath.Focus();
                _logService.LogMessage("検証エラー: フォルダパスが空です。");
                return false;
            }
            if (!Directory.Exists(cmbFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbFolderPath.Focus();
                _logService.LogMessage($"検証エラー: フォルダが存在しません - {cmbFolderPath.Text}");
                return false;
            }
            if (string.IsNullOrWhiteSpace(cmbKeyword.Text))
            {
                MessageBox.Show("検索キーワードを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmbKeyword.Focus();
                _logService.LogMessage("検証エラー: 検索キーワードが空です。");
                return false;
            }

            if (!string.IsNullOrWhiteSpace(txtIgnoreFileSizeMB.Text))
            {
                if (!double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double fileSize) || fileSize < 0)
                {
                    MessageBox.Show("無視ファイルサイズには0以上の数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtIgnoreFileSizeMB.Focus();
                    _logService.LogMessage($"検証エラー: 無視ファイルサイズの入力が無効です - {txtIgnoreFileSizeMB.Text}");
                    return false;
                }
            }
            else // 空の場合は0として扱う（UIにも反映した方が親切かも）
            {
                txtIgnoreFileSizeMB.Text = "0";
                 _logService.LogMessage("無視ファイルサイズが空のため、0として扱います。");
            }

            if (chkRegex.Checked) // 正規表現使用時はコンパイル試行
            {
                try
                {
                    // Regexのコンパイルを試みて、無効ならエラーとする
                    new Regex(cmbKeyword.Text);
                     _logService.LogMessage($"正規表現の事前検証OK: {cmbKeyword.Text}");
                }
                catch (ArgumentException ex)
                {
                    MessageBox.Show($"正規表現のパターンが無効です:\n{ex.Message}", "正規表現エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cmbKeyword.Focus();
                    _logService.LogMessage($"検証エラー: 正規表現が無効です - {cmbKeyword.Text}, Error: {ex.Message}");
                    return false;
                }
            }


            _logService.LogMessage("検索条件の入力検証が正常に完了しました。");
            return true;
        }

        /// <summary>
        /// 検索結果をグリッドにまとめて表示 (リアルタイム表示OFF時や最終表示用)
        /// </summary>
        private void DisplaySearchResults(List<SearchResult> resultsToDisplay)
        {
            if (resultsToDisplay == null)
            {
                _logService.LogMessage("DisplaySearchResults: 表示する結果がnullです。");
                return;
            }
            _logService.LogMessage($"DisplaySearchResults: グリッドをクリアして {resultsToDisplay.Count} 件の結果を表示します。");

            Action updateGridAction = () =>
            {
                grdResults.Rows.Clear(); // まずクリア
                if (resultsToDisplay.Any())
                {
                    AddSearchResultsToGrid(resultsToDisplay); // まとめて追加
                }
                grdResults.Refresh(); // UIの更新を強制
                _logService.LogMessage($"グリッドに {resultsToDisplay.Count} 件の結果を表示しました (DisplaySearchResults)。");
            };

            if (this.InvokeRequired)
            {
                this.Invoke(updateGridAction);
            }
            else
            {
                updateGridAction();
            }
        }


        /// <summary>
        /// 複数の検索結果をグリッドに追加
        /// </summary>
        private void AddSearchResultsToGrid(List<SearchResult> newResults)
        {
            if (newResults == null || !newResults.Any()) return;
            _logService.LogMessage($"AddSearchResultsToGrid: {newResults.Count} 件の結果をグリッドに追加します。");


            Action addRowsAction = () =>
            {
                // DataGridViewRow のリストを作成してから AddRange する方が高速な場合がある
                List<DataGridViewRow> rowsToAdd = new List<DataGridViewRow>();
                foreach (var result in newResults)
                {
                    var row = new DataGridViewRow();
                    row.CreateCells(grdResults); // グリッドのカラム定義に基づいてセルを作成
                    row.Cells[colFilePath.Index].Value = result.FilePath;
                    row.Cells[colFileName.Index].Value = Path.GetFileName(result.FilePath);
                    row.Cells[colSheetName.Index].Value = result.SheetName;
                    row.Cells[colCellPosition.Index].Value = result.CellPosition;
                    row.Cells[colCellValue.Index].Value = TruncateDisplayValue(result.CellValue); // 表示用に切り詰め
                    rowsToAdd.Add(row);
                }

                grdResults.Rows.AddRange(rowsToAdd.ToArray());

                // 最新の行にスクロール (リアルタイム表示時、かつ行が追加された場合のみ)
                if (_isSearching && chkRealTimeDisplay.Checked && grdResults.Rows.Count > 0 && newResults.Any())
                {
                    try
                    {
                         // Ensure будget visible
                        if (grdResults.Rows.Count > 0)
                        {
                            grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logService.LogMessage($"スクロールエラー: {ex.Message}");
                    }
                }
                 // grdResults.Refresh(); // AddRange後、必要ならRefresh。通常は自動で再描画される。
            };


            if (this.InvokeRequired)
            {
                this.Invoke(addRowsAction);
            }
            else
            {
                addRowsAction();
            }
        }

        /// <summary>
        /// グリッド表示用に値を切り詰める
        /// </summary>
        private string TruncateDisplayValue(string value, int maxLength = 500) // 表示用は少し長め
        {
            if (string.IsNullOrEmpty(value)) return value;
            string cleanedValue = value.Replace("\r", "").Replace("\n", "↵"); // 改行は矢印記号に
            return cleanedValue.Length <= maxLength ? cleanedValue : cleanedValue.Substring(0, maxLength) + "...";
        }


        /// <summary>
        /// ステータス表示を更新
        /// </summary>
        private void UpdateStatus(string message)
        {
            _logService.UpdateStatus(message); // LogService経由でUIスレッドを考慮して更新
        }

        /// <summary>
        /// 検索結果をTSVファイルに書き出す
        /// </summary>
        private void WriteResultsToTsv(List<SearchResult> resultsToWrite)
        {
            if (resultsToWrite == null || !resultsToWrite.Any())
            {
                _logService.LogMessage("TSV書き出し: 書き出す結果がありません。");
                return;
            }

            string exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            string fileName = $"SearchResults_{DateTime.Now:yyyyMMdd_HHmmss}.tsv";
            string filePath = Path.Combine(exePath, fileName);
            _logService.LogMessage($"TSVファイル書き出し開始: {filePath} ({resultsToWrite.Count}件)");

            try
            {
                StringBuilder sb = new StringBuilder();
                // ヘッダー行 (DataGridViewのヘッダーに合わせるか、固定にするか)
                // ここでは固定ヘッダーとする
                sb.AppendLine("ファイルパス\tファイル名\tシート名\tセル位置\t検出した値");

                foreach (var result in resultsToWrite)
                {
                    // Path.GetFileName でヌルチェックを追加
                    string resultFileName = string.IsNullOrEmpty(result.FilePath) ? "" : Path.GetFileName(result.FilePath);

                    string tsvRow = string.Join("\t",
                        EscapeTsvField(result.FilePath),
                        EscapeTsvField(resultFileName),
                        EscapeTsvField(result.SheetName),
                        EscapeTsvField(result.CellPosition),
                        EscapeTsvField(result.CellValue) // 生の値を書き出す
                    );
                    sb.AppendLine(tsvRow);
                }

                File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
                _logService.LogMessage($"TSVファイル書き出し完了: {fileName} ({resultsToWrite.Count}件)");
                UpdateStatus($"結果をTSVファイルに書き出しました: {fileName}");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル書き出しエラー: {ex.Message}");
                MessageBox.Show($"TSVファイルへの書き出し中にエラーが発生しました: {ex.Message}", "TSV書き出しエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// TSVフィールドのエスケープ処理
        /// </summary>
        private string EscapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            // タブ、改行、ダブルクォートが含まれる場合はダブルクォートで囲み、中のダブルクォートは2つにする
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
            _logService.LogMessage($"TSVファイル読み込み開始: {filePath}");
            UpdateStatus($"TSVファイル読み込み中: {Path.GetFileName(filePath)} ...");
            Cursor = Cursors.WaitCursor;

            try
            {
                grdResults.Rows.Clear();
                _searchResults.Clear();
                List<SearchResult> loadedResults = new List<SearchResult>();

                // Encoding.UTF8 を指定して読み込む (BOMなしUTF-8にも対応するためには StreamReader のコンストラクタで detectEncodingFromByteOrderMarks: true を使う)
                using (StreamReader reader = new StreamReader(filePath, Encoding.UTF8, true))
                {
                    string line;
                    bool isFirstLine = true; // ヘッダー行スキップ用

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (isFirstLine)
                        {
                            isFirstLine = false;
                            // ヘッダー行の検証やカラムマッピングが必要な場合はここで行う
                            // 今回は固定5列と仮定
                            _logService.LogMessage($"TSVヘッダー行: {line}");
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(line)) continue;

                        string[] fields = ParseTsvLine(line);

                        if (fields.Length >= 5) // 少なくとも5列あることを期待
                        {
                            SearchResult result = new SearchResult
                            {
                                FilePath = UnescapeTsvField(fields[0]),
                                // ファイル名はFilePathから取得するため、TSVの2列目は表示では使わない
                                SheetName = UnescapeTsvField(fields[2]),
                                CellPosition = UnescapeTsvField(fields[3]),
                                CellValue = UnescapeTsvField(fields[4])
                            };
                            loadedResults.Add(result);
                        }
                        else
                        {
                            _logService.LogMessage($"TSV行の形式が不正です (列数不足: {fields.Length}列): {line}");
                        }
                    }
                }

                _searchResults.AddRange(loadedResults); // メインの検索結果リストにも追加
                DisplaySearchResults(_searchResults); // まとめてグリッドに表示

                UpdateStatus($"TSVファイルから {_searchResults.Count} 件の結果を読み込みました。");
                _logService.LogMessage($"TSVファイル読み込み完了: {_searchResults.Count}件");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"TSVファイル読み込みエラー: {ex.ToString()}");
                MessageBox.Show($"TSVファイルの読み込み中にエラーが発生しました: {ex.Message}", "TSV読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("TSVファイルの読み込みに失敗しました。");
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        /// <summary>
        /// TSVの1行をパースする (簡易版、ダブルクォートによるエスケープ対応)
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
                        // エスケープされたダブルクォート ""
                        currentField.Append('"');
                        i++; // 次の文字をスキップ
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
            fields.Add(currentField.ToString()); // 最後のフィールドを追加
            return fields.ToArray();
        }


        /// <summary>
        /// TSVフィールドのアンエスケープ処理
        /// </summary>
        private string UnescapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";

            // 前後のダブルクォートを除去し、中の連続するダブルクォートを一つにする
            if (field.Length >= 2 && field.StartsWith("\"") && field.EndsWith("\""))
            {
                string unescaped = field.Substring(1, field.Length - 2);
                return unescaped.Replace("\"\"", "\"");
            }
            return field;
        }

        /// <summary>
        /// 現在のプロセスのワーキングセットメモリ使用量を取得 (MB単位)
        /// </summary>
        private double GetCurrentMemoryUsageMB()
        {
            Process currentProcess = Process.GetCurrentProcess();
            currentProcess.Refresh();
            return currentProcess.WorkingSet64 / (1024.0 * 1024.0);
        }

    }
}