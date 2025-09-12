using SimpleExcelGrep.Models;
using SimpleExcelGrep.Services;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

namespace SimpleExcelGrep.Forms
{
    /// <summary>
    /// アプリケーションのメインフォーム (コア部分)
    /// </summary>
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

            // UIの初期設定 (グリッドなど)
            InitializeGrid();
        }
        
        /// <summary>
        /// フォーム読み込み時の処理
        /// </summary>
        private void MainForm_Load(object sender, EventArgs e)
        {
            // タイトルにバージョン情報を表示
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            // 設定読み込みより先にログを有効化しておく
            LoadSettings(); 
            _logService.LogMessage("アプリケーション起動");
            _logService.LogEnvironmentInfo();
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

                cmbFolderPath.Items.Clear();
                cmbFolderPath.Items.AddRange(settings.FolderPathHistory.ToArray());
                cmbFolderPath.Text = settings.FolderPath;

                cmbKeyword.Items.Clear();
                cmbKeyword.Items.AddRange(settings.SearchKeywordHistory.ToArray());
                cmbKeyword.Text = settings.SearchKeyword;

                cmbIgnoreKeywords.Items.Clear();
                cmbIgnoreKeywords.Items.AddRange(settings.IgnoreKeywordsHistory.ToArray());
                cmbIgnoreKeywords.Text = settings.IgnoreKeywords;

                chkRegex.Checked = settings.UseRegex;
                chkRealTimeDisplay.Checked = settings.RealTimeDisplay;
                nudParallelism.Value = Math.Min(Math.Max(settings.MaxParallelism, 1), 32);
                chkFirstHitOnly.Checked = settings.FirstHitOnly;
                chkSearchShapes.Checked = settings.SearchShapes;
                txtIgnoreFileSizeMB.Text = settings.IgnoreFileSizeMB.ToString(CultureInfo.InvariantCulture);
                chkCellMode.Checked = settings.CellModeEnabled;
                txtCellAddress.Text = settings.CellAddress;

                // 追加された設定の読み込み
                chkSearchSubDir.Checked = settings.SearchSubDirectories;
                chkEnableLog.Checked = settings.EnableLog;
                _logService.IsLoggingEnabled = settings.EnableLog; // LogServiceの状態も更新
                chkEnableInvisibleSheet.Checked = settings.SearchInvisibleSheets;
                chkDblClickToOpen.Checked = settings.DblClickToOpen;

                txtCellAddress.Enabled = chkCellMode.Checked;
            }
            finally
            {
                _isLoading = false;
            }
        }

        /// <summary>
        /// 現在のUIの状態から設定を保存
        /// </summary>
        private void SaveCurrentSettings()
        {
            if (_isLoading) return;

            double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double ignoreFileSize);

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
                CellModeEnabled = chkCellMode.Checked,
                CellAddress = txtCellAddress.Text,

                // 追加された設定の保存
                SearchSubDirectories = chkSearchSubDir.Checked,
                EnableLog = chkEnableLog.Checked,
                SearchInvisibleSheets = chkEnableInvisibleSheet.Checked,
                DblClickToOpen = chkDblClickToOpen.Checked
            };

            if (!_settingsService.SaveSettings(settings))
            {
                MessageBox.Show("設定の保存に失敗しました。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// UIの検索状態を設定
        /// </summary>
        private void SetSearchingState(bool isSearching)
        {
            _isSearching = isSearching;
            
            // UI要素の有効/無効を一括で切り替え
            var controlsToToggle = new Control[] {
                cmbFolderPath, btnSelectFolder, cmbKeyword, cmbIgnoreKeywords,
                txtIgnoreFileSizeMB, chkRegex, chkSearchShapes, chkFirstHitOnly,
                nudParallelism, btnStartSearch, btnLoadTsv, chkCellMode,
                chkSearchSubDir, chkEnableLog, chkEnableInvisibleSheet, chkCollectStrInShape
            };
            foreach (var control in controlsToToggle)
            {
                control.Enabled = !isSearching;
            }

            txtCellAddress.Enabled = !isSearching && chkCellMode.Checked;
            btnCancelSearch.Enabled = isSearching;

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
        /// ステータス表示を更新
        /// </summary>
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
            _logService.LogMessage($"ステータス更新: {message}", false);
        }
    }
}