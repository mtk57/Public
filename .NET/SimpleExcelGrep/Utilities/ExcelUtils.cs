using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SimpleExcelGrep.Utilities
{
    /// <summary>
    /// Excel関連のユーティリティ機能を提供するクラス
    /// </summary>
    internal static class ExcelUtils
    {
        /// <summary>
        /// 選択行をクリップボードにコピー
        /// </summary>
        public static bool CopySelectedRowsToClipboard(DataGridView grid, Action<string> logCallback)
        {
            if (grid.SelectedRows.Count <= 0) return false;
            
            StringBuilder sb = new StringBuilder();

            // ヘッダー行を追加
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                sb.Append(EscapeForClipboard(grid.Columns[i].HeaderText));
                sb.Append(i == grid.Columns.Count - 1 ? Environment.NewLine : "\t");
            }

            // 選択行を追加（選択順に処理）
            List<DataGridViewRow> selectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                selectedRows.Add(row);
            }

            // インデックスでソート（上から下の順番になるように）
            selectedRows.Sort((x, y) => x.Index.CompareTo(y.Index));

            foreach (DataGridViewRow row in selectedRows)
            {
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    string cellValue = row.Cells[i].Value?.ToString() ?? "";
                    sb.Append(EscapeForClipboard(cellValue));
                    sb.Append(i == grid.Columns.Count - 1 ? Environment.NewLine : "\t");
                }
            }

            try
            {
                DataObject dataObject = new DataObject();
                string textData = sb.ToString();
                dataObject.SetData(DataFormats.UnicodeText, true, textData);
                
                using (MemoryStream stream = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8)) 
                {
                    writer.Write(textData);
                    writer.Flush();
                    stream.Position = 0;
                    dataObject.SetData("Csv", false, stream); 
                }
                
                Clipboard.SetDataObject(dataObject, true);
                logCallback($"{selectedRows.Count}行をクリップボードにコピーしました");
                return true;
            }
            catch (Exception ex)
            {
                logCallback($"クリップボードへのコピーに失敗: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// クリップボード用のテキストをエスケープ
        /// </summary>
        private static string EscapeForClipboard(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            
            bool needsQuoting = text.Contains("\r") || text.Contains("\n") || 
                               text.Contains("\t") || text.Contains("\"");
            
            if (needsQuoting)
            {
                string escapedText = text.Replace("\"", "\"\"");
                return "\"" + escapedText + "\"";
            }
            else
            {
                return text;
            }
        }
    }
}