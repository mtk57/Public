using SimpleExcelGrep.Models;
using SimpleExcelGrep.Utilities;
using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;

namespace SimpleExcelGrep.Forms
{
    /// <summary>
    /// メインフォーム (検索結果グリッド関連の処理部分)
    /// </summary>
    public partial class MainForm
    {
        /// <summary>
        /// グリッドの初期設定
        /// </summary>
        private void InitializeGrid()
        {
            grdResults.MultiSelect = true;
            grdResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            
            var contextMenu = new ContextMenuStrip();
            contextMenu.Items.Add("コピー", null, (s, e) => CopySelectedRowsToClipboard());
            contextMenu.Items.Add("すべて選択", null, (s, e) => grdResults.SelectAll());
            grdResults.ContextMenuStrip = contextMenu;
        }

        /// <summary>
        /// グリッド行ダブルクリック時の処理
        /// </summary>
        private void GrdResults_DoubleClick(object sender, EventArgs e)
        {
            if (grdResults.SelectedRows.Count <= 0) return;

            var selectedRow = grdResults.SelectedRows[0];
            string filePath = selectedRow.Cells["colFilePath"].Value?.ToString();

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (chkDblClickToOpen.Checked)
            {
                if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                {
                    OpenContainingFolder(filePath);
                }
                else
                {
                    string sheetName = selectedRow.Cells["colSheetName"].Value?.ToString();
                    string cellPosition = selectedRow.Cells["colCellPosition"].Value?.ToString();
                    bool isShape = cellPosition == "図形内" || cellPosition == "図形内 (GF)";
                    OpenExcel(filePath, sheetName, isShape ? null : cellPosition);
                }
            }
            else
            {
                OpenContainingFolder(filePath);
            }
        }
        
        /// <summary>
        /// グリッドでのキー押下時の処理 (Ctrl+A, Ctrl+C)
        /// </summary>
        private void GrdResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                grdResults.SelectAll();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectedRowsToClipboard();
                e.Handled = true;
            }
        }

        /// <summary>
        /// ファイルの含まれるフォルダをエクスプローラで開く
        /// </summary>
        private void OpenContainingFolder(string filePath)
        {
            try
            {
                string folderPath = Path.GetDirectoryName(filePath);
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
        /// Excelファイルを開き、指定のセルを選択
        /// </summary>
        private void OpenExcel(string filePath, string sheetName, string cellPosition)
        {
            if (_excelInteropService.OpenExcelFile(filePath, sheetName, cellPosition))
            {
                UpdateStatus($"{Path.GetFileName(filePath)} を開きました。");
            }
            else
            {
                MessageBox.Show("Excelファイルを開けませんでした。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 選択行をクリップボードにコピー
        /// </summary>
        private void CopySelectedRowsToClipboard()
        {
            if (grdResults.SelectedRows.Count == 0) return;
            ExcelUtils.CopySelectedRowsToClipboard(grdResults, message => UpdateStatus(message));
        }
        
        /// <summary>
        /// 検索結果をグリッドに表示 (リスト全体)
        /// </summary>
        private void DisplaySearchResults(List<SearchResult> results)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => DisplaySearchResults(results)));
                return;
            }

            string filePathFilter;
            string fileNameFilter;
            string sheetNameFilter;
            string cellAddressFilter;
            string cellValueFilter;
            GetCurrentGridFilterStrings(out filePathFilter, out fileNameFilter, out sheetNameFilter, out cellAddressFilter, out cellValueFilter);

            grdResults.SuspendLayout();
            grdResults.Rows.Clear();
            var rows = new List<DataGridViewRow>();
            foreach (var result in results)
            {
                if (!MatchesGridFilters(result, filePathFilter, fileNameFilter, sheetNameFilter, cellAddressFilter, cellValueFilter)) continue;

                var row = new DataGridViewRow();
                row.CreateCells(
                    grdResults,
                    result.FilePath,
                    Path.GetFileName(result.FilePath ?? string.Empty),
                    result.SheetName,
                    result.CellPosition,
                    result.CellValue);
                rows.Add(row);
            }
            if (rows.Count > 0)
            {
                grdResults.Rows.AddRange(rows.ToArray());
            }
            grdResults.ResumeLayout();
            _logService.LogMessage($"グリッドに {rows.Count} 件の結果を表示しました (元件数 {results.Count} 件)");
        }

        /// <summary>
        /// 検索結果を1件グリッドに追加
        /// </summary>
        private void AddSearchResultToGrid(SearchResult result)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => AddSearchResultToGrid(result)));
                return;
            }

            string filePathFilter;
            string fileNameFilter;
            string sheetNameFilter;
            string cellAddressFilter;
            string cellValueFilter;
            GetCurrentGridFilterStrings(out filePathFilter, out fileNameFilter, out sheetNameFilter, out cellAddressFilter, out cellValueFilter);

            if (!MatchesGridFilters(result, filePathFilter, fileNameFilter, sheetNameFilter, cellAddressFilter, cellValueFilter)) return;

            grdResults.Rows.Add(
                result.FilePath,
                Path.GetFileName(result.FilePath ?? string.Empty),
                result.SheetName,
                result.CellPosition,
                result.CellValue);
            if (_isSearching && grdResults.Rows.Count > 0)
            {
                grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
            }
        }

        /// <summary>
        /// グリッドフィルタ文字列を初期化
        /// </summary>
        private void ClearGridFilters()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(ClearGridFilters));
                return;
            }

            try
            {
                _isUpdatingGridFilter = true;
                txtFilePathFilter.Clear();
                txtFileNameFilter.Clear();
                txtSheetNameFilter.Clear();
                txtCellAdrFilter.Clear();
                txtCellValueFilter.Clear();
            }
            finally
            {
                _isUpdatingGridFilter = false;
            }

            DisplaySearchResults(_searchResults);
        }

        /// <summary>
        /// 現在のフィルタ文字列を取得
        /// </summary>
        private void GetCurrentGridFilterStrings(out string filePathFilter, out string fileNameFilter, out string sheetNameFilter, out string cellAddressFilter, out string cellValueFilter)
        {
            filePathFilter = NormalizeFilterText(txtFilePathFilter.Text);
            fileNameFilter = NormalizeFilterText(txtFileNameFilter.Text);
            sheetNameFilter = NormalizeFilterText(txtSheetNameFilter.Text);
            cellAddressFilter = NormalizeFilterText(txtCellAdrFilter.Text);
            cellValueFilter = NormalizeFilterText(txtCellValueFilter.Text);
        }

        /// <summary>
        /// 指定の結果が現在のフィルタに一致するか判定
        /// </summary>
        private bool MatchesGridFilters(SearchResult result, string filePathFilter, string fileNameFilter, string sheetNameFilter, string cellAddressFilter, string cellValueFilter)
        {
            string fileName = string.IsNullOrEmpty(result.FilePath) ? string.Empty : Path.GetFileName(result.FilePath);

            return ContainsFilter(result.FilePath, filePathFilter)
                && ContainsFilter(fileName, fileNameFilter)
                && ContainsFilter(result.SheetName, sheetNameFilter)
                && ContainsFilter(result.CellPosition, cellAddressFilter)
                && ContainsFilter(result.CellValue, cellValueFilter);
        }

        /// <summary>
        /// フィルタ文字列を正規化
        /// </summary>
        private static string NormalizeFilterText(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        /// <summary>
        /// 部分一致によるフィルタ判定
        /// </summary>
        private static bool ContainsFilter(string target, string filter)
        {
            if (string.IsNullOrEmpty(filter)) return true;
            if (string.IsNullOrEmpty(target)) return false;
            return target.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
