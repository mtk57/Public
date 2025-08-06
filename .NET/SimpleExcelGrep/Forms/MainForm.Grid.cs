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
            
            grdResults.SuspendLayout();
            grdResults.Rows.Clear();
            var rows = new List<DataGridViewRow>();
            foreach (var result in results)
            {
                var row = new DataGridViewRow();
                row.CreateCells(grdResults, result.FilePath, Path.GetFileName(result.FilePath), result.SheetName, result.CellPosition, result.CellValue);
                rows.Add(row);
            }
            grdResults.Rows.AddRange(rows.ToArray());
            grdResults.ResumeLayout();
            _logService.LogMessage($"グリッドに {results.Count} 件の結果を表示しました");
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

            grdResults.Rows.Add(result.FilePath, Path.GetFileName(result.FilePath), result.SheetName, result.CellPosition, result.CellValue);
            if (_isSearching && grdResults.Rows.Count > 0)
            {
                grdResults.FirstDisplayedScrollingRowIndex = grdResults.Rows.Count - 1;
            }
        }
    }
}