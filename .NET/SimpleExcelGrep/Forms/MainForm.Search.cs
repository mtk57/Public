using SimpleExcelGrep.Models;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleExcelGrep.Forms
{
    /// <summary>
    /// メインフォーム (検索実行と検証ロジック部分)
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// GREP検索モードを実行
        /// </summary>
        private async Task RunGrepSearchModeAsync()
        {
            _logService.LogMessage("GREP検索モードを開始します");
            if (!ValidateGrepInputs()) return;

            await ExecuteSearchAsync(async (pendingResults, token) =>
            {
                var regex = chkRegex.Checked ? new Regex(cmbKeyword.Text, RegexOptions.Compiled | RegexOptions.IgnoreCase) : null;
                var ignoreKeywords = cmbIgnoreKeywords.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(k => k.Trim()).ToList();
                double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var ignoreSize);

                return await _excelSearchService.SearchExcelFilesAsync(cmbFolderPath.Text, cmbKeyword.Text, chkRegex.Checked, regex,
                    ignoreKeywords, chkRealTimeDisplay.Checked, chkSearchShapes.Checked, chkFirstHitOnly.Checked,
                    (int)nudParallelism.Value, ignoreSize, chkSearchSubDir.Checked, chkEnableInvisibleSheet.Checked, 
                    pendingResults, UpdateStatus, token);
            });
        }

        /// <summary>
        /// セル検索モードを実行
        /// </summary>
        private async Task RunCellSearchModeAsync()
        {
            _logService.LogMessage("セル検索モードを開始します");
            var (isValid, addresses) = ValidateCellAddresses();
            if (!isValid) return;

            await ExecuteSearchAsync(async (pendingResults, token) =>
            {
                var ignoreKeywords = cmbIgnoreKeywords.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(k => k.Trim()).ToList();
                double.TryParse(txtIgnoreFileSizeMB.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out var ignoreSize);

                return await _excelSearchService.SearchCellsByAddressAsync(cmbFolderPath.Text, addresses, ignoreKeywords,
                    (int)nudParallelism.Value, ignoreSize, chkSearchSubDir.Checked, chkEnableInvisibleSheet.Checked, 
                    pendingResults, UpdateStatus, token);
            });
        }
        
        /// <summary>
        /// 検索処理の共通実行部分
        /// </summary>
        /// <param name="searchAction">実行する検索タスク</param>
        private async Task ExecuteSearchAsync(Func<ConcurrentQueue<SearchResult>, CancellationToken, Task<List<SearchResult>>> searchAction)
        {
            SetSearchingState(true);
            _cancellationTokenSource = new CancellationTokenSource();
            _searchResults.Clear();
            grdResults.Rows.Clear();
            
            var pendingResults = new ConcurrentQueue<SearchResult>();
            _uiTimer = new System.Windows.Forms.Timer { Interval = 100 };
            StartResultUpdateTimer(_uiTimer, pendingResults, chkRealTimeDisplay.Checked);

            try
            {
                _searchResults = await searchAction(pendingResults, _cancellationTokenSource.Token);
                DisplaySearchResults(_searchResults);
                UpdateStatus($"検索完了: {_searchResults.Count} 件見つかりました");
                if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
            }
            catch (OperationCanceledException)
            {
                UpdateStatus($"検索は中止されました: {_searchResults.Count} 件見つかりました");
                if (!chkRealTimeDisplay.Checked) DisplaySearchResults(_searchResults);
                if (_searchResults.Any()) WriteResultsToTsv(_searchResults);
            }
            catch (Exception ex)
            {
                _logService.LogMessage($"検索エラー: {ex.Message}");
                MessageBox.Show($"検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatus("エラーが発生しました");
                if (!chkRealTimeDisplay.Checked) DisplaySearchResults(_searchResults);
            }
            finally
            {
                _uiTimer?.Stop();
                _uiTimer?.Dispose();
                _uiTimer = null;
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
                if (!isRealTimeDisplay) return;
                
                const int maxUpdatesPerTick = 100;
                for (int i = 0; i < maxUpdatesPerTick && pendingResults.TryDequeue(out var result); i++)
                {
                    _searchResults.Add(result);
                    AddSearchResultToGrid(result);
                }
            };
            timer.Start();
        }

        /// <summary>
        /// フォルダパスの検証
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
            if (chkRegex.Checked)
            {
                try { new Regex(cmbKeyword.Text); }
                catch (ArgumentException ex)
                {
                    MessageBox.Show($"正規表現が無効です: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// セルアドレスの形式を検証
        /// </summary>
        private (bool, string[]) ValidateCellAddresses()
        {
            string[] addresses = txtCellAddress.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                  .Select(a => a.Trim().ToUpper())
                                                  .ToArray();
            bool isValid = addresses.Any() && addresses.All(addr => Regex.IsMatch(addr, @"^[A-Z]{1,3}[1-9][0-9]{0,6}$"));

            if (!isValid)
            {
                MessageBox.Show("セルアドレスの形式が正しくありません。\nA1形式で、カンマ区切りで入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return (false, null);
            }
            return (true, addresses);
        }
    }
}