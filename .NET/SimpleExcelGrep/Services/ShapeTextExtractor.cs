using System;
using System.Collections.Generic;
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

        /// <summary>
        /// WorksheetPartから、グループ化された図形を含め、すべての図形内テキストを抽出する
        /// </summary>
        public IEnumerable<string> ExtractAllTextsFromWorksheetPart(DocumentFormat.OpenXml.Packaging.WorksheetPart worksheetPart)
        {
            var drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null || drawingsPart.WorksheetDrawing == null)
            {
                return Enumerable.Empty<string>();
            }

            var texts = new List<string>();
            // TwoCellAnchor は図形やグループのコンテナ
            foreach (var anchor in drawingsPart.WorksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
            {
                ExtractTextsRecursive(anchor, texts);
            }
            // OneCellAnchor や AbsoluteAnchor も必要に応じて追加
            return texts;
        }

        /// <summary>
        /// 図形コンテナ内を再帰的に探索してテキストを抽出する
        /// </summary>
        private void ExtractTextsRecursive(DocumentFormat.OpenXml.OpenXmlCompositeElement container, List<string> foundTexts)
        {
            // 1. 通常の図形 (Shape) からテキストを抽出
            foreach (var shape in container.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>())
            {
                if (shape.TextBody != null)
                {
                    string text = GetTextFromShapeTextBody(shape.TextBody);
                    if (!string.IsNullOrEmpty(text))
                    {
                        foundTexts.Add(text);
                    }
                }
            }

            // 2. グラフィックフレーム (GraphicFrame) からテキストを抽出 (SmartArtなど)
            foreach (var graphicFrame in container.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>())
            {
                string text = GetTextFromGraphicFrame(graphicFrame);
                if (!string.IsNullOrEmpty(text))
                {
                    foundTexts.Add(text);
                }
            }

            // 3. グループ化図形 (GroupShape) の場合、再帰的に中身を探索
            foreach (var groupShape in container.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>())
            {
                ExtractTextsRecursive(groupShape, foundTexts);
            }
        }
    }
}