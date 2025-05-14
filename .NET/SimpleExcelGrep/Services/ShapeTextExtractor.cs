using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// Excel図形内のテキスト抽出を担当するサービス
    /// </summary>
    internal class ShapeTextExtractor
    {
        /// <summary>
        /// ジェネリックなTextBody要素からテキストを抽出
        /// </summary>
        public string ExtractTextFromTextBodyElement<T>(T textBodyElement) where T : DocumentFormat.OpenXml.OpenXmlCompositeElement
        {
            StringBuilder sb = new StringBuilder();
            if (textBodyElement == null) return string.Empty;

            foreach (var paragraph in textBodyElement.Elements<Paragraph>())
            {
                foreach (var run in paragraph.Elements<Run>())
                {
                    var textElement = run.Elements<Text>().FirstOrDefault();
                    if (textElement != null)
                    {
                        sb.Append(textElement.InnerText);
                    }
                }
                // 段落ごとに改行を挿入（ただし、最後の空の段落は除く）
                if (sb.Length > 0 && !sb.ToString().EndsWith(Environment.NewLine) && paragraph.HasChildren)
                {
                     sb.AppendLine();
                }
            }
            // 末尾の余分な改行を削除
            return sb.ToString().TrimEnd(Environment.NewLine.ToCharArray());
        }

        /// <summary>
        /// ShapeのTextBodyからテキストを抽出
        /// </summary>
        public string GetTextFromShapeTextBody(DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody textBody)
        {
            return ExtractTextFromTextBodyElement(textBody);
        }

        /// <summary>
        /// GraphicFrameからテキストを抽出
        /// </summary>
        public string GetTextFromGraphicFrame(DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame)
        {
            StringBuilder sb = new StringBuilder();
            var graphicData = graphicFrame.Graphic?.GraphicData;
            if (graphicData != null)
            {
                // GraphicData 直下の TextBody (D.TextBody) を探す
                var directTextBodies = graphicData.Elements<DocumentFormat.OpenXml.Drawing.TextBody>();
                foreach(var textBody in directTextBodies)
                {
                    sb.Append(ExtractTextFromTextBodyElement(textBody));
                }

                // さらに深い階層の TextBody (D.TextBody) も探す (例: Chart内など)
                // ただし、無限ループや意図しない要素を拾わないように注意が必要
                // ここでは Descendants を使って簡易的に取得
                var descendantTextBodies = graphicData.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>();
                foreach(var textBody in descendantTextBodies.Except(directTextBodies)) // 重複を避ける
                {
                     if (sb.Length > 0 && textBody.HasChildren) sb.AppendLine(); // 複数のTextBodyが見つかった場合の区切り
                    sb.Append(ExtractTextFromTextBodyElement(textBody));
                }
            }
            return sb.ToString().Trim();
        }
    }
}