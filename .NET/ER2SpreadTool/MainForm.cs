using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using GroupShape = DocumentFormat.OpenXml.Drawing.GroupShape; // OpenXml.Drawing.GroupShape を使用
using Shape = DocumentFormat.OpenXml.Drawing.Shape; // OpenXml.Drawing.Shape を使用
// using TextBody = DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody; // これはShapeの直接の子ではない場合がある
using DrawingTextBody = DocumentFormat.OpenXml.Drawing.TextBody; // Shape内のテキストはこちらを使うことが多い
using Run = DocumentFormat.OpenXml.Drawing.Run;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using Paragraph = DocumentFormat.OpenXml.Drawing.Paragraph;


namespace ER2SpreadTool
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excelファイル (*.xlsx)|*.xlsx";
                openFileDialog.Title = "処理対象のExcelファイルを選択してください";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            // 入力チェック
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("Excelファイルパスを指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSheetName.Text))
            {
                MessageBox.Show("シート名を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("指定されたファイルが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                lblStatus.Text = "処理中...";
                txtResults.Clear();
                Application.DoEvents();

                var results = ProcessExcelFile(txtFilePath.Text, txtSheetName.Text);
                
                if (results.Any()) // results.Count > 0 と同等
                {
                    txtResults.AppendText($"\n抽出結果:\n{FormatResultsForDisplay(results)}");
                    CreateOutputSheet(txtFilePath.Text, "ER図抽出結果", results);
                    lblStatus.Text = $"処理完了 - {results.Count}件のテーブル情報を抽出しました。";
                    MessageBox.Show("処理が完了しました。新規シート「ER図抽出結果」に結果を出力しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.AppendText("\n処理対象の図形が見つかりませんでした。\n");
                    lblStatus.Text = "処理対象の図形が見つかりませんでした。";
                     MessageBox.Show("処理対象の図形が見つかりませんでした。ログを確認してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "処理エラー";
                txtResults.AppendText($"\nエラーが発生しました:\n{ex.ToString()}\n"); // ToString()でスタックトレースも記録
                MessageBox.Show($"処理中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TableInfo> ProcessExcelFile(string filePath, string sheetName)
        {
            var results = new List<TableInfo>();

            using (var document = SpreadsheetDocument.Open(filePath, false)) // 読み取り専用で開く
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    txtResults.AppendText("WorkbookPart が見つかりません。\n");
                    return results;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    throw new ArgumentException($"シート '{sheetName}' が見つかりません。");
                }

                var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (worksheetPart == null)
                {
                     txtResults.AppendText($"WorksheetPart が見つかりません (シートID: {sheet.Id.Value})。\n");
                    return results;
                }
                txtResults.AppendText($"シート '{sheetName}' (ID: {sheet.Id.Value}) を処理中...\n");

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                {
                    txtResults.AppendText("DrawingsPart が見つかりません。図形が存在しない可能性があります。\n");
                    return results;
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;
                if (worksheetDrawing == null)
                {
                     txtResults.AppendText("WorksheetDrawing が見つかりません。\n");
                    return results;
                }
                txtResults.AppendText("WorksheetDrawing を取得しました。\n");

                LogAllShapes(worksheetDrawing); // 詳細なログ取り

                // TwoCellAnchor 要素内のグループ図形を処理
                foreach (var anchor in worksheetDrawing.Elements<TwoCellAnchor>())
                {
                    foreach (var groupShape in anchor.Elements<GroupShape>())
                    {
                        txtResults.AppendText("TwoCellAnchor内のGroupShapeを発見。\n");
                        var tableInfo = ExtractTableInfoFromGroup(groupShape);
                        if (tableInfo != null)
                        {
                            if (!results.Any(r => r.TableName.Equals(tableInfo.TableName, StringComparison.OrdinalIgnoreCase)))
                            {
                                results.Add(tableInfo);
                                txtResults.AppendText($"  => グループからテーブル抽出成功: {tableInfo.TableName} (列数: {tableInfo.Columns.Count})\n");
                            }
                            else
                            {
                                txtResults.AppendText($"  => 重複テーブル名のためスキップ: {tableInfo.TableName}\n");
                            }
                        }
                        else
                        {
                             txtResults.AppendText("  => グループからテーブル情報抽出失敗。\n");
                        }
                    }
                }
            }
            return results;
        }
        
        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== 図形構造の調査 ===\n");
            var childElements = worksheetDrawing.ChildElements.ToList();
            txtResults.AppendText($"WorksheetDrawing の直下の子要素数: {childElements.Count}\n");

            foreach (var element in childElements)
            {
                txtResults.AppendText($"要素タイプ: {element.GetType().Name}\n");
                if (element is TwoCellAnchor anchor)
                {
                    LogTwoCellAnchor(anchor);
                }
                // AbsoluteAnchor や OneCellAnchor も存在しうるが、今回の要件では TwoCellAnchor 内の GroupShape が主
            }
            txtResults.AppendText("=== 調査完了 ===\n\n");
        }

        private void LogTwoCellAnchor(TwoCellAnchor anchor)
        {
            txtResults.AppendText("  TwoCellAnchor 内容:\n");
            foreach (var child in anchor.ChildElements)
            {
                txtResults.AppendText($"    {child.GetType().Name}\n");
                if (child is Shape shape) // 非グループ化Shape
                {
                    LogShapeDetails(shape, "      (単独Shape) ");
                }
                else if (child is GroupShape groupShape)
                {
                    txtResults.AppendText("      GroupShape 内容:\n");
                    var shapesInGroup = groupShape.Elements<Shape>().ToList();
                    txtResults.AppendText($"        グループ内Shape数: {shapesInGroup.Count}\n");
                    foreach (var groupedShape in shapesInGroup)
                    {
                        LogShapeDetails(groupedShape, "          (グループ内Shape) ");
                    }
                }
                // 他にも ConnectionShape などがあり得る
            }
        }

        private void LogShapeDetails(Shape shape, string indent)
        {
            // Shape ID や Name があればログに出す (デバッグ用)
            var nonVisualSpPr = shape.NonVisualShapeProperties;
            if (nonVisualSpPr?.NonVisualDrawingProperties?.Id != null)
            {
                 txtResults.AppendText($"{indent}Shape ID: {nonVisualSpPr.NonVisualDrawingProperties.Id.Value}\n");
            }
            if (nonVisualSpPr?.NonVisualDrawingProperties?.Name != null)
            {
                 txtResults.AppendText($"{indent}Shape Name: {nonVisualSpPr.NonVisualDrawingProperties.Name.Value}\n");
            }

            var drawingTextBody = shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>();// Shape 直下の TextBody (DocumentFormat.OpenXml.Drawing.TextBody)
            if (drawingTextBody != null)
            {
                var extractedText = ExtractTextFromDrawingTextBody(drawingTextBody);
                if (!string.IsNullOrWhiteSpace(extractedText))
                {
                    txtResults.AppendText($"{indent}テキスト: '{extractedText.Replace("\n", "\\n")}'\n");
                }
                else
                {
                    txtResults.AppendText($"{indent}テキストなし (Drawing.TextBody は存在)\n");
                }
            }
            else
            {
                 txtResults.AppendText($"{indent}テキストなし (Drawing.TextBody が null)\n");
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromGroup 開始\n");
            var shapesInGroup = groupShape.Elements<Shape>().ToList(); // Drawing.Shape
            txtResults.AppendText($"    グループ内のShape数: {shapesInGroup.Count}\n");

            // 要件: テキストボックス1(テーブル名)とテキストボックス2(カラム名)がグループ化 [cite: 1]
            if (shapesInGroup.Count != 2)
            {
                txtResults.AppendText($"    グループ内のShape数が2つではありません (実際は{shapesInGroup.Count}個)。処理をスキップします。\n");
                return null;
            }

            var textDataFromShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < shapesInGroup.Count; i++)
            {
                var shape = shapesInGroup[i];
                txtResults.AppendText($"    グループ内Shape {i+1} のテキスト抽出試行...\n");
                var drawingTextBody = shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>();// Drawing.TextBody
                string textContent = string.Empty;
                string[] lines = new string[0];

                if (drawingTextBody != null)
                {
                    textContent = ExtractTextFromDrawingTextBody(drawingTextBody);
                    if (!string.IsNullOrWhiteSpace(textContent))
                    {
                        lines = textContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(l => l.Trim())
                                           .Where(l => !string.IsNullOrEmpty(l))
                                           .ToArray();
                        txtResults.AppendText($"      抽出テキスト: '{textContent.Replace("\n", "\\n")}' (有効行数: {lines.Length})\n");
                    }
                    else
                    {
                        txtResults.AppendText($"      テキストは空または空白でした (Drawing.TextBody は存在)。\n");
                    }
                }
                else
                {
                     txtResults.AppendText($"      Drawing.TextBody が null でした。\n");
                }
                textDataFromShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            
            ShapeTextParseResult tableShapeData = null;
            ShapeTextParseResult columnsShapeData = null;

            // 1つ目のShapeがテーブル名(1行)、2つ目のShapeがカラム名(1行以上)のパターン
            if (textDataFromShapes[0].Lines.Length == 1 && textDataFromShapes[1].Lines.Length >= 1)
            {
                tableShapeData = textDataFromShapes[0];
                columnsShapeData = textDataFromShapes[1];
                txtResults.AppendText("      パターン合致: Shape1=テーブル名, Shape2=カラム名\n");
            }
            // 2つ目のShapeがテーブル名(1行)、1つ目のShapeがカラム名(1行以上)のパターン
            else if (textDataFromShapes[1].Lines.Length == 1 && textDataFromShapes[0].Lines.Length >= 1)
            {
                tableShapeData = textDataFromShapes[1];
                columnsShapeData = textDataFromShapes[0];
                txtResults.AppendText("      パターン合致: Shape2=テーブル名, Shape1=カラム名\n");
            }
            else
            {
                txtResults.AppendText($"    テーブル名(1行)とカラム名(1行以上)の明確な組み合わせが見つかりません。\n");
                txtResults.AppendText($"      Shape1 有効行数: {textDataFromShapes[0].Lines.Length} (元テキスト: '{textDataFromShapes[0].OriginalText.Replace("\n", "\\n")}')\n");
                txtResults.AppendText($"      Shape2 有効行数: {textDataFromShapes[1].Lines.Length} (元テキスト: '{textDataFromShapes[1].OriginalText.Replace("\n", "\\n")}')\n");
                return null;
            }

            string tableName = tableShapeData.Lines[0];
            List<string> columns = columnsShapeData.Lines.Distinct().ToList(); // 重複カラム名は除去

            if (string.IsNullOrWhiteSpace(tableName))
            {
                txtResults.AppendText("    テーブル名が空または空白のため、テーブル情報として不適切です。\n");
                return null;
            }
             if (!columns.Any())
            {
                txtResults.AppendText("    カラムリストが空のため、テーブル情報として不適切です。\n");
                return null;
            }

            txtResults.AppendText($"    テーブル名として確定: '{tableName}'\n");
            txtResults.AppendText($"    カラム名として確定: [{string.Join(", ", columns)}]\n");
            
            return new TableInfo
            {
                TableName = tableName,
                Columns = columns
            };
        }

        // DocumentFormat.OpenXml.Drawing.TextBody からテキストを抽出
        private string ExtractTextFromDrawingTextBody(DrawingTextBody textBody)
        {
            var textParts = new List<string>();
            foreach (var paragraph in textBody.Elements<Paragraph>())
            {
                string paragraphText = "";
                foreach (var run in paragraph.Elements<Run>())
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        if (text.Text != null)
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                 // 改行コードが ParagraphProperties の EndParagraphRunProperties に <a:br/> として存在する場合もあるが、
                 // 通常、Paragraph ごとが改行に対応することが多い。
                 // ここでは、各Paragraphの内容を単純に連結し、後でSplitする。
                textParts.Add(paragraphText);
            }
            // 各Paragraphは通常、Excel内での改行に対応するため、Join時に改行を挟む
            return string.Join("\n", textParts);
        }
        
        // 旧 ExtractTextFromTextBody (Spreadsheet.TextBody用) は今回直接使わないが、参考として残す場合は別に
        // private string ExtractTextFromShape(Shape shape) は不要になり、直接 ExtractTextFromDrawingTextBody を使う

        private string FormatResultsForDisplay(List<TableInfo> results) // txtResults 表示用
        {
            var output = new List<string>();
            output.Add("--- 抽出テーブル一覧 ---");
            foreach (var table in results)
            {
                output.Add($"テーブル名: {table.TableName}");
                output.Add("  カラム名:");
                foreach (var column in table.Columns)
                {
                    output.Add($"    - {column}");
                }
                output.Add(""); // 空行で区切り
            }
            return string.Join("\n", output);
        }

        private void CreateOutputSheet(string filePath, string outputSheetName, List<TableInfo> results)
        {
            using (var document = SpreadsheetDocument.Open(filePath, true)) // 書き込み可能で開く
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null) throw new InvalidOperationException("WorkbookPart is null.");

                // 既存の同名シートを削除
                var existingSheet = workbookPart.Workbook.Descendants<Sheet>()
                                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(outputSheetName, StringComparison.OrdinalIgnoreCase));
                if (existingSheet != null)
                {
                    var sheetId = existingSheet.Id.Value;
                    var existingWorksheetPart = workbookPart.GetPartById(sheetId) as WorksheetPart;
                    if (existingWorksheetPart != null)
                    {
                        workbookPart.DeletePart(existingWorksheetPart);
                    }
                    existingSheet.Remove();
                }

                // 新しいワークシートを作成
                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets == null)
                {
                    sheets = workbookPart.Workbook.AppendChild(new Sheets());
                }

                uint newSheetId = 1;
                if (sheets.Elements<Sheet>().Any())
                {
                    newSheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }
                
                var newSheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(newWorksheetPart),
                    SheetId = newSheetId,
                    Name = outputSheetName
                };
                sheets.Append(newSheet);

                var sheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) // 通常は AddNewPart で作成されるはず
                {
                     sheetData = newWorksheetPart.Worksheet.AppendChild(new SheetData());
                }

                // ヘッダー行
                var headerRow = new Row() { RowIndex = 1U }; // 1-based index
                headerRow.Append(CreateCell("A", 1U, "#"));
                headerRow.Append(CreateCell("B", 1U, "num"));
                headerRow.Append(CreateCell("C", 1U, "テーブル名"));
                headerRow.Append(CreateCell("D", 1U, "カラム名"));
                sheetData.Append(headerRow);

                // データ行
                uint currentRowIndex = 2U;
                int overallCounter = 1;
                foreach (var table in results)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        var dataRow = new Row() { RowIndex = currentRowIndex };
                        dataRow.Append(CreateCell("A", currentRowIndex, overallCounter.ToString()));
                        dataRow.Append(CreateCell("B", currentRowIndex, (i + 1).ToString()));
                        dataRow.Append(CreateCell("C", currentRowIndex, table.TableName));
                        dataRow.Append(CreateCell("D", currentRowIndex, table.Columns[i]));
                        sheetData.Append(dataRow);
                        
                        overallCounter++;
                        currentRowIndex++;
                    }
                }
                workbookPart.Workbook.Save();
            }
        }

        private Cell CreateCell(string columnNamePrefix, uint rowIndex, string value)
        {
            return new Cell()
            {
                CellReference = columnNamePrefix + rowIndex,
                DataType = CellValues.String, // InlineString よりも String + SharedStringTable が推奨される場合もあるが、単純なツールなら InlineString でも可
                                            // DataType = CellValues.InlineString, // これを使う場合は InlineString 要素も必要
                CellValue = new CellValue(value)  // DataType = CellValues.String の場合
                // InlineString = new InlineString(new Text(value)) // DataType = CellValues.InlineString の場合
            };
        }
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }

    // グループ内の各Shapeからテキストをパースした結果を保持する内部クラス
    internal class ShapeTextParseResult
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; } // 空白行を除き、Trimされた有効な行
    }
}