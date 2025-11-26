using System;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using SimpleExcelGrep.Services;

namespace SimpleExcelGrep.Forms
{
    public partial class OtherForm : Form
    {
        private readonly LogService _logService;
        private readonly ExcelModificationService _modificationService;
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning;

        public OtherForm(LogService logService, ExcelModificationService modificationService)
        {
            _logService = logService;
            _modificationService = modificationService;
            InitializeComponent();
            UpdateRunningState(false);
        }

        public void ApplyMainFormState(string folderPath, bool includeSubDirectories, string ignoreFileSizeText, bool includeInvisibleSheets, int parallelism, bool enableLog)
        {
            txtFolderPath.Text = folderPath ?? string.Empty;
            chkSubFolders.Checked = includeSubDirectories;
            txtIgnoreFileSizeMB.Text = ignoreFileSizeText ?? string.Empty;
            chkInvisibleSheets.Checked = includeInvisibleSheets;
            nudParallelism.Value = Math.Max(nudParallelism.Minimum, Math.Min(nudParallelism.Maximum, parallelism));
            chkEnableLog.Checked = enableLog;
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            if (!chkAllShape.Checked && !chkAllFormula.Checked)
            {
                MessageBox.Show("「全ての図」または「全ての数式」にチェックを入れてください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!TryCreateOptions(out var options))
            {
                return;
            }

            if (MessageBox.Show("実行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return;
            }

            _cancellationTokenSource = new CancellationTokenSource();
            UpdateRunningState(true);
            ReportProgress(0, 1, "処理開始...");

            try
            {
                var result = await _modificationService.RunAsync(options, ReportProgress, _cancellationTokenSource.Token);
                var summary = $"処理対象: {result.TotalFiles}件\n変更あり: {result.ModifiedFiles}件\nエラー: {result.ErrorCount}件";
                MessageBox.Show($"処理が完了しました。\n{summary}", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ReportProgress(result.ProcessedFiles, Math.Max(1, result.TotalFiles), "完了");
            }
            catch (OperationCanceledException)
            {
                ReportProgress(0, 1, "処理を中止しました");
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"その他処理で予期せぬエラー: {ex}", force: true);
                MessageBox.Show($"処理中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _cancellationTokenSource?.Dispose();
                _cancellationTokenSource = null;
                UpdateRunningState(false);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            if (_isRunning && _cancellationTokenSource != null)
            {
                _cancellationTokenSource.Cancel();
            }
        }

        private bool TryCreateOptions(out OtherOperationOptions options)
        {
            options = null;

            if (string.IsNullOrWhiteSpace(txtFolderPath.Text))
            {
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            double ignoreFileSize = 0;
            if (!string.IsNullOrWhiteSpace(txtIgnoreFileSizeMB.Text) &&
                !double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out ignoreFileSize))
            {
                MessageBox.Show("無視ファイルサイズは数値で入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            options = new OtherOperationOptions
            {
                FolderPath = txtFolderPath.Text,
                IncludeSubDirectories = chkSubFolders.Checked,
                IgnoreFileSizeMb = ignoreFileSize,
                IncludeInvisibleSheets = chkInvisibleSheets.Checked,
                MaxParallelism = (int)nudParallelism.Value,
                RemoveAllShapes = chkAllShape.Checked,
                RemoveAllFormulas = chkAllFormula.Checked,
                EnableLogOutput = chkEnableLog.Checked
            };

            return true;
        }

        private void UpdateRunningState(bool isRunning)
        {
            _isRunning = isRunning;
            btnRun.Enabled = !isRunning;
            btnCancel.Enabled = isRunning;

            chkAllShape.Enabled = !isRunning;
            chkAllFormula.Enabled = !isRunning;
            txtFolderPath.Enabled = !isRunning;
            chkSubFolders.Enabled = !isRunning;
            txtIgnoreFileSizeMB.Enabled = !isRunning;
            chkInvisibleSheets.Enabled = !isRunning;
            nudParallelism.Enabled = !isRunning;
            chkEnableLog.Enabled = !isRunning;
        }

        private void ReportProgress(int processed, int total, string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ReportProgress(processed, total, message)));
                return;
            }

            progressBar1.Maximum = Math.Max(1, total);
            progressBar1.Value = Math.Max(0, Math.Min(processed, progressBar1.Maximum));
            lblProgress.Text = message;
        }
    }
}
