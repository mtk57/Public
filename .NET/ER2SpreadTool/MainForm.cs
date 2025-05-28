using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml; // For OpenXmlCompositeElement
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
// Explicitly use full namespaces for Drawing elements to avoid ambiguity
// Aliases can be used if preferred, but full names are clearer here.
// using Drawing = DocumentFormat.OpenXml.Drawing;
// using DrawingSpreadsheet = DocumentFormat.OpenXml.Drawing.Spreadsheet;

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
                
                if (results.Any())
                {
                    txtResults.AppendText($"\n抽出結果:\n{FormatResultsForDisplay(results)}");
                    CreateOutputSheet(txtFilePath.Text, "ER図抽出結果", results); // Ensure this method works fine
                    lblStatus.Text = $"処理完了 - {results.Count}件のテーブル情報を抽出しました。";
                    MessageBox.Show("処理が完了しました。新規シート「ER図抽出結果」に結果を出力しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.AppendText("\n処理対象の図形が見つかりませんでした。\nデバッグ情報を確認してください。");
                    lblStatus.Text = "処理対象の図形が見つかりませんでした。";
                    MessageBox.Show("処理対象の図形が見つかりませんでした。ログを確認してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "処理エラー";
                txtResults.AppendText($"\nエラーが発生しました:\n{ex.ToString()}\n");
                MessageBox.Show($"処理中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TableInfo> ProcessExcelFile(string filePath, string sheetName)
        {
            var results = new List<TableInfo>();
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    txtResults.AppendText("WorkbookPart が見つかりません。\n");
                    return results;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null) throw new ArgumentException($"シート '{sheetName}' が見つかりません。");

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
                LogAllShapes(worksheetDrawing);

                foreach (var anchor in worksheetDrawing.Elements<TwoCellAnchor>())
                {
                    // Iterate through DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape (xdr:grpSp) elements
                    foreach (var xdrGroupShape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>())
                    {
                        txtResults.AppendText("Spreadsheet.GroupShape (xdr:grpSp) をTwoCellAnchor内で発見。\n");
                        var tableInfo = ExtractTableInfoFromSpreadsheetGroup(xdrGroupShape);
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
                             txtResults.AppendText("  => Spreadsheet.GroupShapeからテーブル情報抽出失敗。\n");
                        }
                    }
                }
            }
            return results;
        }
        
        private TableInfo ExtractTableInfoFromSpreadsheetGroup(DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromSpreadsheetGroup (xdr:grpSp) 開始\n");

            // Scenario 1: The xdr:grpSp contains a nested a:grpSp (Drawing.GroupShape)
            var nestedDrawingGroup = xdrGroupShape.GetFirstChild<DocumentFormat.OpenXml.Drawing.GroupShape>();
            if (nestedDrawingGroup != null)
            {
                txtResults.AppendText("    ネストされた Drawing.GroupShape (a:grpSp) を発見。これを処理します。\n");
                return ExtractTableInfoFromDrawingGroupShape(nestedDrawingGroup);
            }

            // Scenario 2: The xdr:grpSp directly contains two xdr:sp (Spreadsheet.Shape) elements
            var xdrShapesInGroup = xdrGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
            txtResults.AppendText($"    xdr:grpSp内のSpreadsheet.Shape (xdr:sp) 数: {xdrShapesInGroup.Count}\n");

            if (xdrShapesInGroup.Count != 2)
            {
                txtResults.AppendText($"    xdr:grpSp内のxdr:sp数が2つではありません。現ロジックでは処理不可。\n");
                return null;
            }

            var textDataFromXdrShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < xdrShapesInGroup.Count; i++)
            {
                var xdrShape = xdrShapesInGroup[i];
                txtResults.AppendText($"    xdr:sp {i + 1} のテキスト抽出試行...\n");
                string textContent = string.Empty;
                
                // Text in an xdr:sp is usually in its xdr:txBody (Spreadsheet.TextBody)
                var spreadsheetTextBody = xdrShape.TextBody; // This is DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody
                if (spreadsheetTextBody != null)
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) を発見。\n");
                    textContent = ExtractTextFromSpreadsheetTextBody(spreadsheetTextBody);
                }
                else
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) が見つかりません。\n");
                }

                string[] lines = new string[0];
                if (!string.IsNullOrWhiteSpace(textContent))
                {
                    lines = textContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(l => l.Trim())
                                       .Where(l => !string.IsNullOrEmpty(l))
                                       .ToArray();
                    txtResults.AppendText($"      抽出テキスト: '{textContent.Replace("\n", "\\n")}' (有効行数: {lines.Length})\n");
                }
                textDataFromXdrShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromXdrShapes, "xdr:spベース");
        }

        private TableInfo ExtractTableInfoFromDrawingGroupShape(DocumentFormat.OpenXml.Drawing.GroupShape drawingGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromDrawingGroupShape (a:grpSp) 開始\n");
            var shapesInDrawingGroup = drawingGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Shape>().ToList(); // These are a:sp
            txtResults.AppendText($"    a:grpSp内のDrawing.Shape (a:sp) 数: {shapesInDrawingGroup.Count}\n");

            if (shapesInDrawingGroup.Count != 2)
            {
                txtResults.AppendText($"    a:grpSp内のa:sp数が2つではありません。\n");
                return null;
            }

            var textDataFromDrawingShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < shapesInDrawingGroup.Count; i++)
            {
                var drawingShape = shapesInDrawingGroup[i]; // This is a:sp
                txtResults.AppendText($"    a:sp {i + 1} のテキスト抽出試行...\n");
                var drawingTextBody = drawingShape.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>(); // a:sp -> a:txBody
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
                }
                else
                {
                    txtResults.AppendText($"      a:sp {i + 1} に Drawing.TextBody (a:txBody) が見つかりません。\n");
                }
                textDataFromDrawingShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromDrawingShapes, "a:spベース");
        }
        
        private TableInfo CreateTableInfoFromTextData(List<ShapeTextParseResult> textDataList, string sourceDescription)
        {
            if (textDataList.Count != 2) {
                 txtResults.AppendText($"    {sourceDescription}: テキスト情報を2つ取得できませんでした (実際は{textDataList.Count}個)。\n");
                return null;
            }

            ShapeTextParseResult tableShapeData = null;
            ShapeTextParseResult columnsShapeData = null;

            if (textDataList[0].Lines.Length == 1 && textDataList[1].Lines.Length >= 1)
            {
                tableShapeData = textDataList[0];
                columnsShapeData = textDataList[1];
            }
            else if (textDataList[1].Lines.Length == 1 && textDataList[0].Lines.Length >= 1)
            {
                tableShapeData = textDataList[1];
                columnsShapeData = textDataList[0];
            }
            else
            {
                txtResults.AppendText($"    テーブル名(1行)とカラム名(1行以上)の明確な組み合わせが見つかりません ({sourceDescription})。\n");
                txtResults.AppendText($"      Shape1 有効行数: {textDataList[0].Lines.Length} (元テキスト: '{textDataList[0].OriginalText.Replace("\n", "\\n")}')\n");
                txtResults.AppendText($"      Shape2 有効行数: {textDataList[1].Lines.Length} (元テキスト: '{textDataList[1].OriginalText.Replace("\n", "\\n")}')\n");
                return null;
            }

            // Ensure that the identified table name shape actually has content
            if (tableShapeData.Lines.Length == 0 || string.IsNullOrWhiteSpace(tableShapeData.Lines[0]))
            {
                txtResults.AppendText($"    {sourceDescription}: 抽出されたテーブル名が空です。\n");
                return null;
            }
            // Ensure that the identified columns shape actually has content
             if (columnsShapeData.Lines.Length == 0)
            {
                txtResults.AppendText($"    {sourceDescription}: 抽出されたカラムリストが空です。\n");
                return null;
            }


            string tableName = tableShapeData.Lines[0];
            List<string> columns = columnsShapeData.Lines.Distinct().ToList();
            
            txtResults.AppendText($"    {sourceDescription}: テーブル名として確定: '{tableName}'\n");
            txtResults.AppendText($"    {sourceDescription}: カラム名として確定: [{string.Join(", ", columns)}]\n");

            return new TableInfo { TableName = tableName, Columns = columns };
        }

        private string ExtractTextFromDrawingTextBody(DocumentFormat.OpenXml.Drawing.TextBody drawingTextBody) // For a:txBody
        {
            var textParts = new List<string>();
            foreach (var paragraph in drawingTextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>()) // a:p
            {
                string paragraphText = "";
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>()) // a:r
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Drawing.Text>()) // a:t
                    {
                        if (text.Text != null)
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                textParts.Add(paragraphText);
            }
            return string.Join("\n", textParts);
        }

        private string ExtractTextFromSpreadsheetTextBody(DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody spreadsheetTextBody) // For xdr:txBody
        {
            var textParts = new List<string>();
            // xdr:txBody also contains a:p elements
            foreach (var paragraph in spreadsheetTextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                string paragraphText = "";
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                {
                    foreach (var text in run.Elements<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        if (text.Text != null)
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                textParts.Add(paragraphText);
            }
            return string.Join("\n", textParts);
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== 図形構造の調査 ===\n");
            var childElements = worksheetDrawing.ChildElements.ToList();
            txtResults.AppendText($"WorksheetDrawing の直下の子要素数: {childElements.Count}\n");

            foreach (var element in childElements)
            {
                txtResults.AppendText($"要素タイプ: {element.GetType().FullName}\n"); // Use FullName for clarity
                if (element is TwoCellAnchor anchor)
                {
                    LogTwoCellAnchor(anchor);
                }
            }
            txtResults.AppendText("=== 調査完了 ===\n\n");
        }

        private void LogTwoCellAnchor(TwoCellAnchor anchor)
        {
            txtResults.AppendText("  TwoCellAnchor 内容:\n");
            foreach (var child in anchor.ChildElements)
            {
                txtResults.AppendText($"    子要素タイプ: {child.GetType().FullName}\n");
                if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape xdrShape)
                {
                    txtResults.AppendText("      詳細: Spreadsheet.Shape (xdr:sp)\n");
                    var spTextBody = xdrShape.TextBody;
                    if(spTextBody != null) {
                        txtResults.AppendText($"        xdr:sp Text: '{ExtractTextFromSpreadsheetTextBody(spTextBody).Replace("\n", "\\n")}'\n");
                    }
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
                {
                    txtResults.AppendText("      詳細: Spreadsheet.GroupShape (xdr:grpSp)\n");
                    foreach(var innerElement in xdrGroupShape.Elements<OpenXmlCompositeElement>()){
                        txtResults.AppendText($"        xdr:grpSp 内の子要素: {innerElement.GetType().FullName}\n");
                        if(innerElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape innerXdrShape) {
                             var innerSpTextBody = innerXdrShape.TextBody;
                             if(innerSpTextBody != null) {
                                txtResults.AppendText($"          xdr:sp (内) Text: '{ExtractTextFromSpreadsheetTextBody(innerSpTextBody).Replace("\n", "\\n")}'\n");
                            }
                        } else if (innerElement is DocumentFormat.OpenXml.Drawing.GroupShape innerDrawingGroupShape) { // a:grpSp
                             txtResults.AppendText($"          Drawing.GroupShape (a:grpSp) (内)\n");
                        }
                    }
                }
            }
        }
        
        private string FormatResultsForDisplay(List<TableInfo> results)
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
                output.Add("");
            }
            return string.Join("\n", output);
        }

        private void CreateOutputSheet(string filePath, string outputSheetName, List<TableInfo> results)
        {
             // Using "true" to open for read/write
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null) throw new InvalidOperationException("WorkbookPart is null.");

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
                     workbookPart.Workbook.Save(); // Save after removing sheet part and sheet element
                }

                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets == null) sheets = workbookPart.Workbook.AppendChild(new Sheets());

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
                
                var headerRow = new Row() { RowIndex = 1U };
                headerRow.Append(CreateCell("A", 1U, "#"));
                headerRow.Append(CreateCell("B", 1U, "num"));
                headerRow.Append(CreateCell("C", 1U, "テーブル名"));
                headerRow.Append(CreateCell("D", 1U, "カラム名"));
                sheetData.Append(headerRow);

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
                newWorksheetPart.Worksheet.Save(); // Save the worksheet part
                workbookPart.Workbook.Save(); // Save the workbook
            }
        }

        private Cell CreateCell(string columnNamePrefix, uint rowIndex, string value)
        {
            return new Cell()
            {
                CellReference = columnNamePrefix + rowIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(value)
            };
        }
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }

    internal class ShapeTextParseResult
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; }
    }
}