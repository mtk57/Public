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

            // cmbFolderPath でドラッグアンドドロップを有効化
            cmbFolderPath.AllowDrop = true;
            cmbFolderPath.DragEnter += CmbFolderPath_DragEnter;
            cmbFolderPath.DragDrop += CmbFolderPath_DragDrop;

            // 設定関連のUIイベント
            var controls = new Control[] {
                chkRealTimeDisplay, nudParallelism, chkFirstHitOnly, chkSearchShapes,
                txtIgnoreFileSizeMB, chkCellMode, txtCellAddress,
                chkSearchSubDir, chkEnableInvisibleSheet, chkCollectStrInShape // chkEnableLogは別途処理
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

            var filterTextBoxes = new TextBox[]
            {
                txtFilePathFilter,
                txtFileNameFilter,
                txtSheetNameFilter,
                txtCellAdrFilter,
                txtCellValueFilter
            };
            foreach (var filterTextBox in filterTextBoxes)
            {
                filterTextBox.TextChanged += GridFilterTextBox_TextChanged;
            }
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

            if (chkCollectStrInShape.Checked)
            {
                await RunCollectShapeTextModeAsync();
            }
            else if (chkCellMode.Checked && !string.IsNullOrWhiteSpace(txtCellAddress.Text))
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
        /// グリッドフィルタテキスト変更時の処理
        /// </summary>
        private void GridFilterTextBox_TextChanged(object sender, EventArgs e)
        {
            if (_isUpdatingGridFilter) return;
            DisplaySearchResults(_searchResults);
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

        /// <summary>
        /// フォルダパスのコンボボックスにオブジェクトがドラッグされた際の処理
        /// </summary>
        private void CmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグされているデータがファイル/フォルダであるかを確認
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // ドロップを許可するエフェクトを表示
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                // ファイル/フォルダ以外は許可しない
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// フォルダパスのコンボボックスにオブジェクトがドロップされた際の処理
        /// </summary>
        private void CmbFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            // ドロップされたファイルのパスリストを取得
            string[] paths = (string[])e.Data.GetData(DataFormats.FileDrop);

            if (paths != null && paths.Length > 0)
            {
                string droppedPath = paths[0]; // 複数ドロップされた場合は最初の項目を使用
                string targetDirectory = null;

                // ドロップされたのがフォルダかどうかを判定
                if (Directory.Exists(droppedPath))
                {
                    targetDirectory = droppedPath;
                }
                // ドロップされたのがファイルなら、そのファイルを含むディレクトリを取得
                else if (File.Exists(droppedPath))
                {
                    targetDirectory = Path.GetDirectoryName(droppedPath);
                }

                // 有効なディレクトリパスが取得できた場合
                if (!string.IsNullOrEmpty(targetDirectory))
                {
                    // テキストを更新し、履歴に追加して設定を保存
                    cmbFolderPath.Text = targetDirectory;
                    _settingsService.AddToComboBoxHistory(cmbFolderPath, targetDirectory);
                    SaveCurrentSettings();
                }
            }
        }
    }
}
