using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
        private string _folderPath = string.Empty;
        private bool _includeSubDirectories;
        private double _ignoreFileSizeMb;
        private bool _includeInvisibleSheets;
        private int _parallelism;
        private bool _enableLog;

        public OtherForm(LogService logService, ExcelModificationService modificationService)
        {
            _logService = logService;
            _modificationService = modificationService;
            InitializeComponent();
            UpdateRunningState(false);
            UpdateDeleteByKeywordControls(false);
        }

        public void ApplyMainFormState(string folderPath, bool includeSubDirectories, string ignoreFileSizeText, bool includeInvisibleSheets, int parallelism, bool enableLog)
        {
            _folderPath = folderPath ?? string.Empty;
            _includeSubDirectories = includeSubDirectories;
            if (!double.TryParse(ignoreFileSizeText, NumberStyles.Any, CultureInfo.InvariantCulture, out _ignoreFileSizeMb))
            {
                _ignoreFileSizeMb = 0;
            }
            _includeInvisibleSheets = includeInvisibleSheets;
            _parallelism = parallelism;
            _enableLog = enableLog;
        }

        private async void BtnRun_Click(object sender, EventArgs e)
        {
            if (!chkAllShape.Checked && !chkAllFormula.Checked && !chkEnableDeleteByKeyword.Checked)
            {
                MessageBox.Show("処理対象にチェックを入れてください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            if (string.IsNullOrWhiteSpace(_folderPath))
            {
                MessageBox.Show("フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(_folderPath))
            {
                MessageBox.Show("指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!TryParseDeleteByKeywordOptions(out var deleteByKeyword))
            {
                return false;
            }

            options = new OtherOperationOptions
            {
                FolderPath = _folderPath,
                IncludeSubDirectories = _includeSubDirectories,
                IgnoreFileSizeMb = _ignoreFileSizeMb,
                IncludeInvisibleSheets = _includeInvisibleSheets,
                MaxParallelism = _parallelism,
                RemoveAllShapes = chkAllShape.Checked,
                RemoveAllFormulas = chkAllFormula.Checked,
                EnableLogOutput = _enableLog,
                DeleteByKeywordEnabled = deleteByKeyword.Enabled,
                DeleteByKeywordDirectionIsRow = deleteByKeyword.DirectionIsRow,
                DeleteByKeywordTargetAll = deleteByKeyword.TargetAll,
                DeleteByKeywordTargets = deleteByKeyword.TargetIndices,
                DeleteByKeywordKeywords = deleteByKeyword.Keywords,
                DeleteByKeywordFullMatch = deleteByKeyword.FullMatch,
                DeleteByKeywordCaseSensitive = deleteByKeyword.CaseSensitive,
                DeleteByKeywordWidthSensitive = deleteByKeyword.WidthSensitive
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
            chkEnableDeleteByKeyword.Enabled = !isRunning;
            UpdateDeleteByKeywordControls(isRunning);
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

        private void ChkEnableDeleteByKeyword_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDeleteByKeywordControls(_isRunning);
        }

        private void ChkAll_CheckedChanged(object sender, EventArgs e)
        {
            UpdateDeleteByKeywordControls(_isRunning);
        }

        private void UpdateDeleteByKeywordControls(bool isRunning)
        {
            var enabled = !isRunning && chkEnableDeleteByKeyword.Checked;
            rbtnRow.Enabled = enabled;
            rbtnClm.Enabled = enabled;
            chkAll.Enabled = enabled;
            chkFullMatch.Enabled = enabled;
            chkCaseSensitive.Enabled = enabled;
            chkWidthSensitive.Enabled = enabled;
            txtTargetNum.Enabled = enabled && !chkAll.Checked;
            txtKeyword.Enabled = enabled;
            label1.Enabled = enabled;
            label2.Enabled = enabled;
        }

        private bool TryParseDeleteByKeywordOptions(out DeleteByKeywordOption parsed)
        {
            parsed = null;

            if (!chkEnableDeleteByKeyword.Checked)
            {
                parsed = new DeleteByKeywordOption();
                return true;
            }

            var directionIsRow = rbtnRow.Checked;
            var targetAll = chkAll.Checked;

            if (!TryParseTargets(txtTargetNum.Text, directionIsRow, targetAll, out var targets, out var errorMessage))
            {
                MessageBox.Show(errorMessage, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            var keywords = (txtKeyword.Text ?? string.Empty)
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(k => k.Trim())
                .Where(k => !string.IsNullOrEmpty(k))
                .ToArray();

            if (keywords.Length == 0)
            {
                MessageBox.Show("キーワードを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            parsed = new DeleteByKeywordOption
            {
                Enabled = true,
                DirectionIsRow = directionIsRow,
                TargetAll = targetAll,
                TargetIndices = targets,
                Keywords = keywords,
                FullMatch = chkFullMatch.Checked,
                CaseSensitive = chkCaseSensitive.Checked,
                WidthSensitive = chkWidthSensitive.Checked
            };

            return true;
        }

        private bool TryParseTargets(string input, bool directionIsRow, bool targetAll, out HashSet<int> targets, out string errorMessage)
        {
            targets = new HashSet<int>();
            errorMessage = string.Empty;

            if (targetAll && string.IsNullOrWhiteSpace(input))
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(input))
            {
                errorMessage = "行番号（列番号）を入力してください。";
                return false;
            }

            var tokens = input
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToArray();

            if (tokens.Length == 0)
            {
                errorMessage = "行番号（列番号）を入力してください。";
                return false;
            }

            var hasLetter = tokens.Any(t => t.Any(char.IsLetter));
            var hasDigit = tokens.Any(t => t.Any(char.IsDigit));

            if (hasLetter && hasDigit)
            {
                errorMessage = "行番号と列番号が混在しています。";
                return false;
            }

            var inputIsRowNumber = !hasLetter;

            if (directionIsRow && !inputIsRowNumber)
            {
                errorMessage = "「行」が選択されています。行番号を入力してください。";
                return false;
            }

            if (!directionIsRow && inputIsRowNumber)
            {
                errorMessage = "「列」が選択されています。列番号を入力してください。";
                return false;
            }

            foreach (var token in tokens)
            {
                if (!TryParseRangeToken(token, inputIsRowNumber, targets, out errorMessage))
                {
                    return false;
                }
            }

            if (targets.Count == 0)
            {
                errorMessage = "有効な行番号（列番号）が指定されていません。";
                return false;
            }

            return true;
        }

        private bool TryParseRangeToken(string token, bool isRowNumber, HashSet<int> targets, out string errorMessage)
        {
            errorMessage = string.Empty;
            var rangeParts = token.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries)
                                  .Select(p => p.Trim())
                                  .ToArray();

            if (rangeParts.Length == 1)
            {
                if (!TryParseSingleValue(rangeParts[0], isRowNumber, out var value, out errorMessage))
                {
                    return false;
                }

                targets.Add(value);
                return true;
            }

            if (rangeParts.Length == 2)
            {
                if (!TryParseSingleValue(rangeParts[0], isRowNumber, out var start, out errorMessage) ||
                    !TryParseSingleValue(rangeParts[1], isRowNumber, out var end, out errorMessage))
                {
                    return false;
                }

                if (start > end)
                {
                    errorMessage = "範囲指定が不正です。開始と終了を確認してください。";
                    return false;
                }

                for (int i = start; i <= end; i++)
                {
                    targets.Add(i);
                }

                return true;
            }

            errorMessage = "範囲指定が不正です。";
            return false;
        }

        private bool TryParseSingleValue(string value, bool isRowNumber, out int parsed, out string errorMessage)
        {
            parsed = 0;
            errorMessage = string.Empty;

            if (isRowNumber)
            {
                if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out parsed) || parsed <= 0)
                {
                    errorMessage = "行番号は1以上の整数で指定してください。";
                    return false;
                }

                return true;
            }

            parsed = ColumnNameToIndex(value);
            if (parsed <= 0)
            {
                errorMessage = "列番号はAから始まる英字で指定してください。";
                return false;
            }

            return true;
        }

        private int ColumnNameToIndex(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName)) return 0;

            int index = 0;
            foreach (var ch in columnName.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') return 0;
                index = index * 26 + (ch - 'A' + 1);
            }

            return index;
        }

        private class DeleteByKeywordOption
        {
            public bool Enabled { get; set; }
            public bool DirectionIsRow { get; set; }
            public bool TargetAll { get; set; }
            public HashSet<int> TargetIndices { get; set; } = new HashSet<int>();
            public string[] Keywords { get; set; } = Array.Empty<string>();
            public bool FullMatch { get; set; }
            public bool CaseSensitive { get; set; }
            public bool WidthSensitive { get; set; }
        }
    }
}
