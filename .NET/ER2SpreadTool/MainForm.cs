using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ER2SpreadTool
{
    public partial class MainForm : Form
    {
        public MainForm ()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click ( object sender, EventArgs e )
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

        private void btnProcess_Click ( object sender, EventArgs e )
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
                Application.DoEvents();

                var results = ProcessExcelFile(txtFilePath.Text, txtSheetName.Text);
                
                txtResults.Text = FormatResults(results);
                lblStatus.Text = $"処理完了 - {results.Count}件のテーブル情報を抽出しました。";

                // 新規シートに出力
                CreateOutputSheet(txtFilePath.Text, results);
                
                MessageBox.Show("処理が完了しました。新規シート「ER図抽出結果」に結果を出力しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "処理エラー";
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
                var drawingsPart = worksheetPart.DrawingsPart;

                if (drawingsPart != null)
                {
                    var worksheetDrawing = drawingsPart.WorksheetDrawing;
                    
                    // グループ化された図形を検索
                    var groupShapes = worksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>();
                    
                    foreach (var groupShape in groupShapes)
                    {
                        var tableInfo = ExtractTableInfoFromGroup(groupShape);
                        if (tableInfo != null)
                        {
                            results.Add(tableInfo);
                        }
                    }
                }
            }

            return results;
        }

        private TableInfo ExtractTableInfoFromGroup(DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape groupShape)
        {
            string tableName = null;
            var columns = new List<string>();

            // グループ内のシェイプを取得
            var shapes = groupShape.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>();

            foreach (var shape in shapes)
            {
                var textBody = shape.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().FirstOrDefault();
                if (textBody != null)
                {
                    var text = ExtractTextFromTextBody(textBody);
                    
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // テキストが単一行の場合はテーブル名、複数行の場合はカラム名とみなす
                        var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(l => l.Trim())
                                       .Where(l => !string.IsNullOrEmpty(l))
                                       .ToArray();

                        if (lines.Length == 1)
                        {
                            // テーブル名候補
                            if (string.IsNullOrEmpty(tableName))
                            {
                                tableName = lines[0];
                            }
                        }
                        else if (lines.Length > 1)
                        {
                            // カラム名候補
                            columns.AddRange(lines);
                        }
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

        private string ExtractTextFromTextBody(DocumentFormat.OpenXml.Drawing.TextBody textBody)
        {
            var textParts = new List<string>();
            
            foreach (var paragraph in textBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>())
            {
                var paragraphText = "";
                foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>())
                {
                    var text = run.Elements<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault();
                    if (text != null)
                    {
                        paragraphText += text.Text;
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
                InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(value))
            };
        }
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }
}
