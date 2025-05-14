using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcelGrep.Models
{
    /// <summary>
    /// 検索結果を格納するクラス
    /// </summary>
    internal class SearchResult
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }
        public string CellPosition { get; set; } // セル位置または "図形内" など
        public string CellValue { get; set; }    // セルの値または図形内のテキスト
    }
}
