using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using GroupShape = DocumentFormat.OpenXml.Drawing.GroupShape;
using Shape = DocumentFormat.OpenXml.Drawing.Shape;
using TextBody = DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using Text = DocumentFormat.OpenXml.Drawing.Text;

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
                
                if (results.Count > 0)
                {
                    txtResults.Text = FormatResults(results);
                    CreateOutputSheet(txtFilePath.Text, results);
                    lblStatus.Text = $"処理完了 - {results.Count}件のテーブル情報を抽出しました。";
                    MessageBox.Show("処理が完了しました。新規シート「ER図抽出結果」に結果を出力しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.Text = "図形が見つかりませんでした。\nデバッグ情報を確認してください。";
                    lblStatus.Text = "図形が見つかりませんでした。";
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "処理エラー";
                txtResults.AppendText($"エラーが発生しました:\n{ex.Message}\n\n");
                MessageBox.Show($"処理中にエラーが発生しました:\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TableInfo> ProcessExcelFile(string filePath, string sheetName)
        {
            var results = new List<TableInfo>();

            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var workbookPart = document.WorkbookPart;
                var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                    .FirstOrDefault(s => s.Name == sheetName);

                if (sheet == null)
                {
                    throw new ArgumentException($"シート '{sheetName}' が見つかりません。");
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                txtResults.AppendText($"シート '{sheetName}' を処理中...\n");

                // DrawingsPart の確認
                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                {
                    txtResults.AppendText("DrawingsPart が見つかりません。図形が存在しない可能性があります。\n");
                    return results;
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;
                txtResults.AppendText("WorksheetDrawing を取得しました。\n");

                // すべての図形要素を詳細に調査
                LogAllShapes(worksheetDrawing);

                // グループ化された図形を処理
                var groupShapes = worksheetDrawing.Descendants<GroupShape>();
                txtResults.AppendText($"グループシェイプ数: {groupShapes.Count()}\n");

                foreach (var groupShape in groupShapes)
                {
                    var tableInfo = ExtractTableInfoFromGroup(groupShape);
                    if (tableInfo != null)
                    {
                        results.Add(tableInfo);
                    }
                }

                // 個別の図形も処理
                var shapes = worksheetDrawing.Descendants<Shape>();
                txtResults.AppendText($"個別シェイプ数: {shapes.Count()}\n");

                ProcessIndividualShapes(shapes, results);
            }

            return results;
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== 図形構造の調査 ===\n");

            // すべての子要素を列挙
            var allElements = worksheetDrawing.ChildElements;
            txtResults.AppendText($"子要素数: {allElements.Count}\n");

            foreach (var element in allElements)
            {
                txtResults.AppendText($"要素タイプ: {element.GetType().Name}\n");

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
                txtResults.AppendText($"    {child.GetType().Name}\n");

                if (child is Shape shape)
                {
                    LogShapeDetails(shape, "      ");
                }
                else if (child is GroupShape groupShape)
                {
                    txtResults.AppendText("      GroupShape 内容:\n");
                    foreach (var groupChild in groupShape.ChildElements)
                    {
                        txtResults.AppendText($"        {groupChild.GetType().Name}\n");
                        if (groupChild is Shape groupedShape)
                        {
                            LogShapeDetails(groupedShape, "          ");
                        }
                    }
                }
            }
        }

        private void LogShapeDetails(Shape shape, string indent)
        {
            txtResults.AppendText($"{indent}Shape 詳細:\n");

            foreach (var shapeChild in shape.ChildElements)
            {
                txtResults.AppendText($"{indent}  {shapeChild.GetType().Name}\n");

                if (shapeChild is TextBody textBody)
                {
                    var extractedText = ExtractTextFromTextBody(textBody);
                    txtResults.AppendText($"{indent}    テキスト: '{extractedText}'\n");
                }
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            string tableName = null;
            var columns = new List<string>();

            var shapes = groupShape.Descendants<Shape>();
            txtResults.AppendText($"グループ内のシェイプ数: {shapes.Count()}\n");

            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                
                if (!string.IsNullOrWhiteSpace(text))
                {
                    txtResults.AppendText($"抽出テキスト: '{text}'\n");

                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrEmpty(l))
                                   .ToArray();

                    if (lines.Length == 1)
                    {
                        if (string.IsNullOrEmpty(tableName))
                        {
                            tableName = lines[0];
                            txtResults.AppendText($"テーブル名として認識: '{tableName}'\n");
                        }
                    }
                    else if (lines.Length > 1)
                    {
                        columns.AddRange(lines);
                        txtResults.AppendText($"カラム名として認識: [{string.Join(", ", lines)}]\n");
                    }
                }
            }

            if (!string.IsNullOrEmpty(tableName) && columns.Any())
            {
                return new TableInfo
                {
                    TableName = tableName,
                    Columns = columns.Distinct().ToList()
                };
            }

            return null;
        }

        private void ProcessIndividualShapes(IEnumerable<Shape> shapes, List<TableInfo> results)
        {
            var textList = new List<string>();

            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    textList.Add(text);
                    txtResults.AppendText($"個別シェイプテキスト: '{text}'\n");
                }
            }

            // 簡易的な組み合わせロジック（隣接するテキストを組み合わせ）
            for (int i = 0; i < textList.Count - 1; i++)
            {
                var current = textList[i].Trim();
                var next = textList[i + 1].Trim();

                var currentLines = current.Split('\n').Length;
                var nextLines = next.Split('\n').Length;

                if (currentLines == 1 && nextLines > 1)
                {
                    var columns = next.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(l => l.Trim())
                                     .Where(l => !string.IsNullOrEmpty(l))
                                     .ToList();

                    results.Add(new TableInfo
                    {
                        TableName = current,
                        Columns = columns
                    });

                    txtResults.AppendText($"組み合わせ認識 - テーブル: '{current}', カラム: [{string.Join(", ", columns)}]\n");
                }
            }
        }

        private string ExtractTextFromShape(Shape shape)
        {
            var textParts = new List<string>();

            // TextBody からの抽出
            var textBodies = shape.Descendants<TextBody>();
            foreach (var textBody in textBodies)
            {
                var text = ExtractTextFromTextBody(textBody);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    textParts.Add(text);
                }
            }

            return string.Join("\n", textParts);
        }

        private string ExtractTextFromTextBody(TextBody textBody)
        {
            var textParts = new List<string>();
            
            var paragraphs = textBody.Elements<Paragraph>();
            foreach (var paragraph in paragraphs)
            {
                var paragraphText = "";
                var runs = paragraph.Elements<Run>();
                foreach (var run in runs)
                {
                    var textElements = run.Elements<Text>();
                    foreach (var text in textElements)
                    {
                        if (!string.IsNullOrEmpty(text.Text))
                        {
                            paragraphText += text.Text;
                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(paragraphText))
                {
                    textParts.Add(paragraphText.Trim());
                }
            }

            return string.Join("\n", textParts);
        }

        private string FormatResults(List<TableInfo> results)
        {
            var output = new List<string>();
            output.Add("# | num | テーブル名 | カラム名");
            output.Add("--|-----|------------|----------");

            int counter = 1;
            foreach (var table in results)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    output.Add($"{counter} | {i + 1} | {table.TableName} | {table.Columns[i]}");
                    counter++;
                }
            }

            return string.Join("\n", output);
        }

        private void CreateOutputSheet(string filePath, List<TableInfo> results)
        {
            using (var document = SpreadsheetDocument.Open(filePath, true))
            {
                var workbookPart = document.WorkbookPart;
                var sheets = workbookPart.Workbook.Sheets;

                // 既存の同名シートを削除
                var existingSheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == "ER図抽出結果");
                if (existingSheet != null)
                {
                    var existingWorksheetPart = (WorksheetPart)workbookPart.GetPartById(existingSheet.Id);
                    workbookPart.DeletePart(existingWorksheetPart);
                    existingSheet.Remove();
                }

                // 新しいワークシートを作成
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet(new SheetData());
                worksheetPart.Worksheet = worksheet;

                // シートを追加
                var outputSheetName = "ER図抽出結果";
                var newSheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)(sheets.Count() + 1),
                    Name = outputSheetName
                };
                sheets.Append(newSheet);

                // データを書き込み
                var sheetData = worksheet.GetFirstChild<SheetData>();
                
                // ヘッダー行
                var headerRow = new Row() { RowIndex = 1 };
                headerRow.Append(CreateCell("A", 1, "#"));
                headerRow.Append(CreateCell("B", 1, "num"));
                headerRow.Append(CreateCell("C", 1, "テーブル名"));
                headerRow.Append(CreateCell("D", 1, "カラム名"));
                sheetData.Append(headerRow);

                // データ行
                uint rowIndex = 2;
                int counter = 1;
                
                foreach (var table in results)
                {
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        var dataRow = new Row() { RowIndex = rowIndex };
                        dataRow.Append(CreateCell("A", rowIndex, counter.ToString()));
                        dataRow.Append(CreateCell("B", rowIndex, (i + 1).ToString()));
                        dataRow.Append(CreateCell("C", rowIndex, table.TableName));
                        dataRow.Append(CreateCell("D", rowIndex, table.Columns[i]));
                        sheetData.Append(dataRow);
                        
                        counter++;
                        rowIndex++;
                    }
                }

                workbookPart.Workbook.Save();
            }
        }

        private Cell CreateCell(string columnName, uint rowIndex, string value)
        {
            return new Cell()
            {
                CellReference = columnName + rowIndex,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value))
            };
        }
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }

}