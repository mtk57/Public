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
                var twoCellAnchors = worksheetDrawing.Elements<TwoCellAnchor>();
                txtResults.AppendText($"TwoCellAnchor数: {twoCellAnchors.Count()}\n");

                foreach (var anchor in twoCellAnchors)
                {
                    // グループ化された図形を探す
                    var groupShapes = anchor.Elements<GroupShape>();
                    foreach (var groupShape in groupShapes)
                    {
                        var tableInfo = ExtractTableInfoFromGroup(groupShape);
                        if (tableInfo != null)
                        {
                            results.Add(tableInfo);
                            txtResults.AppendText($"グループからテーブル抽出: {tableInfo.TableName} (列数: {tableInfo.Columns.Count})\n");
                        }
                    }

                    // 個別の図形も処理
                    var shapes = anchor.Elements<Shape>();
                    if (shapes.Any())
                    {
                        ProcessShapesInAnchor(shapes, results);
                    }
                }

                // 一次元配列として処理されている図形も確認
                var allShapes = worksheetDrawing.Descendants<Shape>();
                txtResults.AppendText($"全図形数: {allShapes.Count()}\n");
                
                ProcessIndividualShapes(allShapes, results);
            }

            return results;
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== 図形構造の調査 ===\n");

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
                    var groupShapes = groupShape.Elements<Shape>();
                    txtResults.AppendText($"        グループ内Shape数: {groupShapes.Count()}\n");
                    
                    foreach (var groupedShape in groupShapes)
                    {
                        LogShapeDetails(groupedShape, "        ");
                    }
                }
            }
        }

        private void LogShapeDetails(Shape shape, string indent)
        {
            foreach (var shapeChild in shape.ChildElements)
            {
                if (shapeChild is TextBody textBody)
                {
                    var extractedText = ExtractTextFromTextBody(textBody);
                    if (!string.IsNullOrWhiteSpace(extractedText))
                    {
                        txtResults.AppendText($"{indent}テキスト: '{extractedText}'\n");
                    }
                }
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            string tableName = null;
            var columns = new List<string>();

            // グループ内のすべてのShape要素を取得
            var shapes = groupShape.Elements<Shape>();
            txtResults.AppendText($"グループ内のシェイプ数: {shapes.Count()}\n");

            var allTexts = new List<ShapeTextInfo>();

            // 各シェイプからテキストを抽出
            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrEmpty(l))
                                   .ToArray();

                    allTexts.Add(new ShapeTextInfo
                    {
                        OriginalText = text,
                        Lines = lines
                    });

                    txtResults.AppendText($"抽出テキスト: '{text}' (行数: {lines.Length})\n");
                }
            }

            // テーブル名とカラムを識別
            // 1行のテキストをテーブル名、複数行のテキストをカラムとして扱う
            foreach (var textInfo in allTexts)
            {
                if (textInfo.Lines.Length == 1)
                {
                    // 1行 = テーブル名の可能性
                    if (string.IsNullOrEmpty(tableName))
                    {
                        tableName = textInfo.Lines[0];
                        txtResults.AppendText($"テーブル名として認識: '{tableName}'\n");
                    }
                }
                else if (textInfo.Lines.Length > 1)
                {
                    // 複数行 = カラム名の可能性
                    columns.AddRange(textInfo.Lines);
                    txtResults.AppendText($"カラム名として認識: [{string.Join(", ", textInfo.Lines)}]\n");
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

        private void ProcessShapesInAnchor(IEnumerable<Shape> shapes, List<TableInfo> results)
        {
            var textInfos = new List<ShapeTextInfo>();

            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrEmpty(l))
                                   .ToArray();

                    textInfos.Add(new ShapeTextInfo
                    {
                        OriginalText = text,
                        Lines = lines
                    });
                }
            }

            // 隣接するテキストボックスを組み合わせる
            CombineAdjacentTexts(textInfos, results);
        }

        private void ProcessIndividualShapes(IEnumerable<Shape> shapes, List<TableInfo> results)
        {
            var textInfos = new List<ShapeTextInfo>();

            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrEmpty(l))
                                   .ToArray();

                    textInfos.Add(new ShapeTextInfo
                    {
                        OriginalText = text,
                        Lines = lines
                    });

                    txtResults.AppendText($"個別シェイプテキスト: '{text}'\n");
                }
            }

            CombineAdjacentTexts(textInfos, results);
        }

        private void CombineAdjacentTexts(List<ShapeTextInfo> textInfos, List<TableInfo> results)
        {
            // 簡単な組み合わせロジック：1行のテキストと複数行のテキストを組み合わせる
            for (int i = 0; i < textInfos.Count; i++)
            {
                var current = textInfos[i];
                
                if (current.Lines.Length == 1) // テーブル名候補
                {
                    // 次の要素を探してカラム候補を見つける
                    for (int j = i + 1; j < textInfos.Count; j++)
                    {
                        var next = textInfos[j];
                        if (next.Lines.Length > 1) // カラム候補
                        {
                            var tableInfo = new TableInfo
                            {
                                TableName = current.Lines[0],
                                Columns = next.Lines.ToList()
                            };

                            // 重複チェック
                            if (!results.Any(r => r.TableName == tableInfo.TableName))
                            {
                                results.Add(tableInfo);
                                txtResults.AppendText($"組み合わせ認識 - テーブル: '{tableInfo.TableName}', カラム: [{string.Join(", ", tableInfo.Columns)}]\n");
                            }
                            break;
                        }
                    }
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

    public class ShapeTextInfo
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; }
    }
}