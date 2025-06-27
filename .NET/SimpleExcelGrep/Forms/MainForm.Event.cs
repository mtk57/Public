using System;
using System.IO;
using System.Windows.Forms;

namespace SimpleExcelGrep.Forms
{
    /// <summary>
    /// メインフォーム (UIイベントハンドラ部分)
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// イベントハンドラを登録
        /// </summary>
        private void RegisterEventHandlers()
        {
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;

            btnSelectFolder.Click += BtnSelectFolder_Click;
            btnStartSearch.Click += BtnStartSearch_Click;
            btnCancelSearch.Click += BtnCancelSearch_Click;
            btnLoadTsv.Click += BtnLoadTsv_Click;

            grdResults.DoubleClick += GrdResults_DoubleClick;
            grdResults.KeyDown += GrdResults_KeyDown;
            
            cmbKeyword.KeyDown += CmbKeyword_KeyDown;

            // 設定関連のUIイベント
            var controls = new Control[] {
                chkRealTimeDisplay, nudParallelism, chkFirstHitOnly, chkSearchShapes,
                txtIgnoreFileSizeMB, chkCellMode, txtCellAddress,
                chkSearchSubDir, chkEnableInvisibleSheet // chkEnableLogは別途処理
            };
            foreach (var control in controls)
            {
                if (control is CheckBox chk) chk.CheckedChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
                if (control is TextBox txt) txt.TextChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
                if (control is NumericUpDown nud) nud.ValueChanged += (s, e) => { if (!_isLoading) SaveCurrentSettings(); };
            }

            // ログ有効化チェックボックスは、LogServiceの状態も変更する必要があるため、専用のハンドラを登録
            chkEnableLog.CheckedChanged += ChkEnableLog_CheckedChanged;
            
            chkCellMode.CheckedChanged += ChkCellMode_CheckedChanged;
        }

        /// <summary>
        /// 検索開始ボタンクリック時の処理
        /// </summary>
        private async void BtnStartSearch_Click(object sender, EventArgs e)
        {
            if (!ValidateFolderPath()) return;

            _settingsService.AddToComboBoxHistory(cmbFolderPath, cmbFolderPath.Text);
            _settingsService.AddToComboBoxHistory(cmbKeyword, cmbKeyword.Text);
            _settingsService.AddToComboBoxHistory(cmbIgnoreKeywords, cmbIgnoreKeywords.Text);
            SaveCurrentSettings();

            if (chkCellMode.Checked && !string.IsNullOrWhiteSpace(txtCellAddress.Text))
            {
                await RunCellSearchModeAsync();
            }
            else
            {
                await RunGrepSearchModeAsync();
            }
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
        /// フォルダ選択ボタンクリック時の処理
        /// </summary>
        private void BtnSelectFolder_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "検索するフォルダを選択してください";
                dialog.ShowNewFolderButton = false;
                if (!string.IsNullOrEmpty(cmbFolderPath.Text) && Directory.Exists(cmbFolderPath.Text))
                {
                    dialog.SelectedPath = cmbFolderPath.Text;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _settingsService.AddToComboBoxHistory(cmbFolderPath, dialog.SelectedPath);
                    SaveCurrentSettings();
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
                BtnStartSearch_Click(this, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        /// <summary>
        /// セルモードチェックボックス変更時の処理
        /// </summary>
        private void ChkCellMode_CheckedChanged(object sender, EventArgs e)
        {
            txtCellAddress.Enabled = chkCellMode.Checked;
            if (!_isLoading) SaveCurrentSettings();
        }

        /// <summary>
        /// ログ出力チェックボックス変更時の処理
        /// </summary>
        private void ChkEnableLog_CheckedChanged(object sender, EventArgs e)
        {
            if (!_isLoading)
            {
                _logService.IsLoggingEnabled = chkEnableLog.Checked;
                SaveCurrentSettings();
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
    }
}