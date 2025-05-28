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
                var twoCellAnchors = worksheetDrawing.Elements<TwoCellAnchor>();
                txtResults.AppendText($"TwoCellAnchor��: {twoCellAnchors.Count()}\n");

                foreach (var anchor in twoCellAnchors)
                {
                    // �O���[�v�����ꂽ�}�`��T��
                    var groupShapes = anchor.Elements<GroupShape>();
                    foreach (var groupShape in groupShapes)
                    {
                        var tableInfo = ExtractTableInfoFromGroup(groupShape);
                        if (tableInfo != null)
                        {
                            results.Add(tableInfo);
                            txtResults.AppendText($"�O���[�v����e�[�u�����o: {tableInfo.TableName} (��: {tableInfo.Columns.Count})\n");
                        }
                    }

                    // �ʂ̐}�`������
                    var shapes = anchor.Elements<Shape>();
                    if (shapes.Any())
                    {
                        ProcessShapesInAnchor(shapes, results);
                    }
                }

                // �ꎟ���z��Ƃ��ď�������Ă���}�`���m�F
                var allShapes = worksheetDrawing.Descendants<Shape>();
                txtResults.AppendText($"�S�}�`��: {allShapes.Count()}\n");
                
                ProcessIndividualShapes(allShapes, results);
            }

            return results;
        }

        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== �}�`�\���̒��� ===\n");

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
                    var groupShapes = groupShape.Elements<Shape>();
                    txtResults.AppendText($"        �O���[�v��Shape��: {groupShapes.Count()}\n");
                    
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
                        txtResults.AppendText($"{indent}�e�L�X�g: '{extractedText}'\n");
                    }
                }
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            string tableName = null;
            var columns = new List<string>();

            // �O���[�v���̂��ׂĂ�Shape�v�f���擾
            var shapes = groupShape.Elements<Shape>();
            txtResults.AppendText($"�O���[�v���̃V�F�C�v��: {shapes.Count()}\n");

            var allTexts = new List<ShapeTextInfo>();

            // �e�V�F�C�v����e�L�X�g�𒊏o
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

                    txtResults.AppendText($"���o�e�L�X�g: '{text}' (�s��: {lines.Length})\n");
                }
            }

            // �e�[�u�����ƃJ����������
            // 1�s�̃e�L�X�g���e�[�u�����A�����s�̃e�L�X�g���J�����Ƃ��Ĉ���
            foreach (var textInfo in allTexts)
            {
                if (textInfo.Lines.Length == 1)
                {
                    // 1�s = �e�[�u�����̉\��
                    if (string.IsNullOrEmpty(tableName))
                    {
                        tableName = textInfo.Lines[0];
                        txtResults.AppendText($"�e�[�u�����Ƃ��ĔF��: '{tableName}'\n");
                    }
                }
                else if (textInfo.Lines.Length > 1)
                {
                    // �����s = �J�������̉\��
                    columns.AddRange(textInfo.Lines);
                    txtResults.AppendText($"�J�������Ƃ��ĔF��: [{string.Join(", ", textInfo.Lines)}]\n");
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

            // �אڂ���e�L�X�g�{�b�N�X��g�ݍ��킹��
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

                    txtResults.AppendText($"�ʃV�F�C�v�e�L�X�g: '{text}'\n");
                }
            }

            CombineAdjacentTexts(textInfos, results);
        }

        private void CombineAdjacentTexts(List<ShapeTextInfo> textInfos, List<TableInfo> results)
        {
            // �ȒP�ȑg�ݍ��킹���W�b�N�F1�s�̃e�L�X�g�ƕ����s�̃e�L�X�g��g�ݍ��킹��
            for (int i = 0; i < textInfos.Count; i++)
            {
                var current = textInfos[i];
                
                if (current.Lines.Length == 1) // �e�[�u�������
                {
                    // ���̗v�f��T���ăJ��������������
                    for (int j = i + 1; j < textInfos.Count; j++)
                    {
                        var next = textInfos[j];
                        if (next.Lines.Length > 1) // �J�������
                        {
                            var tableInfo = new TableInfo
                            {
                                TableName = current.Lines[0],
                                Columns = next.Lines.ToList()
                            };

                            // �d���`�F�b�N
                            if (!results.Any(r => r.TableName == tableInfo.TableName))
                            {
                                results.Add(tableInfo);
                                txtResults.AppendText($"�g�ݍ��킹�F�� - �e�[�u��: '{tableInfo.TableName}', �J����: [{string.Join(", ", tableInfo.Columns)}]\n");
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

    public class ShapeTextInfo
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; }
    }
}