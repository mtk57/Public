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
                openFileDialog.Filter = "Excel�t�@�C�� (*.xlsx)|*.xlsx";
                openFileDialog.Title = "�����Ώۂ�Excel�t�@�C����I�����Ă�������";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            // ���̓`�F�b�N
            if (string.IsNullOrWhiteSpace(txtFilePath.Text))
            {
                MessageBox.Show("Excel�t�@�C���p�X���w�肵�Ă��������B", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtSheetName.Text))
            {
                MessageBox.Show("�V�[�g������͂��Ă��������B", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("�w�肳�ꂽ�t�@�C�������݂��܂���B", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                lblStatus.Text = "������...";
                txtResults.Clear();
                Application.DoEvents();

                var results = ProcessExcelFile(txtFilePath.Text, txtSheetName.Text);
                
                if (results.Count > 0)
                {
                    txtResults.Text = FormatResults(results);
                    CreateOutputSheet(txtFilePath.Text, results);
                    lblStatus.Text = $"�������� - {results.Count}���̃e�[�u�����𒊏o���܂����B";
                    MessageBox.Show("�������������܂����B�V�K�V�[�g�uER�}���o���ʁv�Ɍ��ʂ��o�͂��܂����B", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.Text = "�}�`��������܂���ł����B\n�f�o�b�O�����m�F���Ă��������B";
                    lblStatus.Text = "�}�`��������܂���ł����B";
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "�����G���[";
                txtResults.AppendText($"�G���[���������܂���:\n{ex.Message}\n\n");
                MessageBox.Show($"�������ɃG���[���������܂���:\n{ex.Message}", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    throw new ArgumentException($"�V�[�g '{sheetName}' ��������܂���B");
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                txtResults.AppendText($"�V�[�g '{sheetName}' ��������...\n");

                // DrawingsPart �̊m�F
                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                {
                    txtResults.AppendText("DrawingsPart ��������܂���B�}�`�����݂��Ȃ��\��������܂��B\n");
                    return results;
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;
                txtResults.AppendText("WorksheetDrawing ���擾���܂����B\n");

                // ���ׂĂ̐}�`�v�f���ڍׂɒ���
                LogAllShapes(worksheetDrawing);

                // �O���[�v�����ꂽ�}�`������
                var groupShapes = worksheetDrawing.Descendants<GroupShape>();
                txtResults.AppendText($"�O���[�v�V�F�C�v��: {groupShapes.Count()}\n");

                foreach (var groupShape in groupShapes)
                {
                    var tableInfo = ExtractTableInfoFromGroup(groupShape);
                    if (tableInfo != null)
                    {
                        results.Add(tableInfo);
                    }
                }

                // �ʂ̐}�`������
                var shapes = worksheetDrawing.Descendants<Shape>();
                txtResults.AppendText($"�ʃV�F�C�v��: {shapes.Count()}\n");

                ProcessIndividualShapes(shapes, results);
            }

            return results;
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== �}�`�\���̒��� ===\n");

            // ���ׂĂ̎q�v�f���
            var allElements = worksheetDrawing.ChildElements;
            txtResults.AppendText($"�q�v�f��: {allElements.Count}\n");

            foreach (var element in allElements)
            {
                txtResults.AppendText($"�v�f�^�C�v: {element.GetType().Name}\n");

                if (element is TwoCellAnchor anchor)
                {
                    LogTwoCellAnchor(anchor);
                }
            }

            txtResults.AppendText("=== �������� ===\n\n");
        }

        private void LogTwoCellAnchor(TwoCellAnchor anchor)
        {
            txtResults.AppendText("  TwoCellAnchor ���e:\n");

            foreach (var child in anchor.ChildElements)
            {
                txtResults.AppendText($"    {child.GetType().Name}\n");

                if (child is Shape shape)
                {
                    LogShapeDetails(shape, "      ");
                }
                else if (child is GroupShape groupShape)
                {
                    txtResults.AppendText("      GroupShape ���e:\n");
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
            txtResults.AppendText($"{indent}Shape �ڍ�:\n");

            foreach (var shapeChild in shape.ChildElements)
            {
                txtResults.AppendText($"{indent}  {shapeChild.GetType().Name}\n");

                if (shapeChild is TextBody textBody)
                {
                    var extractedText = ExtractTextFromTextBody(textBody);
                    txtResults.AppendText($"{indent}    �e�L�X�g: '{extractedText}'\n");
                }
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            string tableName = null;
            var columns = new List<string>();

            var shapes = groupShape.Descendants<Shape>();
            txtResults.AppendText($"�O���[�v���̃V�F�C�v��: {shapes.Count()}\n");

            foreach (var shape in shapes)
            {
                var text = ExtractTextFromShape(shape);
                
                if (!string.IsNullOrWhiteSpace(text))
                {
                    txtResults.AppendText($"���o�e�L�X�g: '{text}'\n");

                    var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(l => l.Trim())
                                   .Where(l => !string.IsNullOrEmpty(l))
                                   .ToArray();

                    if (lines.Length == 1)
                    {
                        if (string.IsNullOrEmpty(tableName))
                        {
                            tableName = lines[0];
                            txtResults.AppendText($"�e�[�u�����Ƃ��ĔF��: '{tableName}'\n");
                        }
                    }
                    else if (lines.Length > 1)
                    {
                        columns.AddRange(lines);
                        txtResults.AppendText($"�J�������Ƃ��ĔF��: [{string.Join(", ", lines)}]\n");
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
                    txtResults.AppendText($"�ʃV�F�C�v�e�L�X�g: '{text}'\n");
                }
            }

            // �ȈՓI�ȑg�ݍ��킹���W�b�N�i�אڂ���e�L�X�g��g�ݍ��킹�j
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

                    txtResults.AppendText($"�g�ݍ��킹�F�� - �e�[�u��: '{current}', �J����: [{string.Join(", ", columns)}]\n");
                }
            }
        }

        private string ExtractTextFromShape(Shape shape)
        {
            var textParts = new List<string>();

            // TextBody ����̒��o
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
            output.Add("# | num | �e�[�u���� | �J������");
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

                // �����̓����V�[�g���폜
                var existingSheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == "ER�}���o����");
                if (existingSheet != null)
                {
                    var existingWorksheetPart = (WorksheetPart)workbookPart.GetPartById(existingSheet.Id);
                    workbookPart.DeletePart(existingWorksheetPart);
                    existingSheet.Remove();
                }

                // �V�������[�N�V�[�g���쐬
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet(new SheetData());
                worksheetPart.Worksheet = worksheet;

                // �V�[�g��ǉ�
                var outputSheetName = "ER�}���o����";
                var newSheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = (uint)(sheets.Count() + 1),
                    Name = outputSheetName
                };
                sheets.Append(newSheet);

                // �f�[�^����������
                var sheetData = worksheet.GetFirstChild<SheetData>();
                
                // �w�b�_�[�s
                var headerRow = new Row() { RowIndex = 1 };
                headerRow.Append(CreateCell("A", 1, "#"));
                headerRow.Append(CreateCell("B", 1, "num"));
                headerRow.Append(CreateCell("C", 1, "�e�[�u����"));
                headerRow.Append(CreateCell("D", 1, "�J������"));
                sheetData.Append(headerRow);

                // �f�[�^�s
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