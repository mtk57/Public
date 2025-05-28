using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using GroupShape = DocumentFormat.OpenXml.Drawing.GroupShape; // OpenXml.Drawing.GroupShape ���g�p
using Shape = DocumentFormat.OpenXml.Drawing.Shape; // OpenXml.Drawing.Shape ���g�p
// using TextBody = DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody; // �����Shape�̒��ڂ̎q�ł͂Ȃ��ꍇ������
using DrawingTextBody = DocumentFormat.OpenXml.Drawing.TextBody; // Shape���̃e�L�X�g�͂�������g�����Ƃ�����
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
                
                if (results.Any()) // results.Count > 0 �Ɠ���
                {
                    txtResults.AppendText($"\n���o����:\n{FormatResultsForDisplay(results)}");
                    CreateOutputSheet(txtFilePath.Text, "ER�}���o����", results);
                    lblStatus.Text = $"�������� - {results.Count}���̃e�[�u�����𒊏o���܂����B";
                    MessageBox.Show("�������������܂����B�V�K�V�[�g�uER�}���o���ʁv�Ɍ��ʂ��o�͂��܂����B", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.AppendText("\n�����Ώۂ̐}�`��������܂���ł����B\n");
                    lblStatus.Text = "�����Ώۂ̐}�`��������܂���ł����B";
                     MessageBox.Show("�����Ώۂ̐}�`��������܂���ł����B���O���m�F���Ă��������B", "���", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "�����G���[";
                txtResults.AppendText($"\n�G���[���������܂���:\n{ex.ToString()}\n"); // ToString()�ŃX�^�b�N�g���[�X���L�^
                MessageBox.Show($"�������ɃG���[���������܂���:\n{ex.Message}", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<TableInfo> ProcessExcelFile(string filePath, string sheetName)
        {
            var results = new List<TableInfo>();

            using (var document = SpreadsheetDocument.Open(filePath, false)) // �ǂݎ���p�ŊJ��
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    txtResults.AppendText("WorkbookPart ��������܂���B\n");
                    return results;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (sheet == null)
                {
                    throw new ArgumentException($"�V�[�g '{sheetName}' ��������܂���B");
                }

                var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                if (worksheetPart == null)
                {
                     txtResults.AppendText($"WorksheetPart ��������܂��� (�V�[�gID: {sheet.Id.Value})�B\n");
                    return results;
                }
                txtResults.AppendText($"�V�[�g '{sheetName}' (ID: {sheet.Id.Value}) ��������...\n");

                var drawingsPart = worksheetPart.DrawingsPart;
                if (drawingsPart == null)
                {
                    txtResults.AppendText("DrawingsPart ��������܂���B�}�`�����݂��Ȃ��\��������܂��B\n");
                    return results;
                }

                var worksheetDrawing = drawingsPart.WorksheetDrawing;
                if (worksheetDrawing == null)
                {
                     txtResults.AppendText("WorksheetDrawing ��������܂���B\n");
                    return results;
                }
                txtResults.AppendText("WorksheetDrawing ���擾���܂����B\n");

                LogAllShapes(worksheetDrawing); // �ڍׂȃ��O���

                // TwoCellAnchor �v�f���̃O���[�v�}�`������
                foreach (var anchor in worksheetDrawing.Elements<TwoCellAnchor>())
                {
                    foreach (var groupShape in anchor.Elements<GroupShape>())
                    {
                        txtResults.AppendText("TwoCellAnchor����GroupShape�𔭌��B\n");
                        var tableInfo = ExtractTableInfoFromGroup(groupShape);
                        if (tableInfo != null)
                        {
                            if (!results.Any(r => r.TableName.Equals(tableInfo.TableName, StringComparison.OrdinalIgnoreCase)))
                            {
                                results.Add(tableInfo);
                                txtResults.AppendText($"  => �O���[�v����e�[�u�����o����: {tableInfo.TableName} (��: {tableInfo.Columns.Count})\n");
                            }
                            else
                            {
                                txtResults.AppendText($"  => �d���e�[�u�����̂��߃X�L�b�v: {tableInfo.TableName}\n");
                            }
                        }
                        else
                        {
                             txtResults.AppendText("  => �O���[�v����e�[�u����񒊏o���s�B\n");
                        }
                    }
                }
            }
            return results;
        }
        
        private void LogAllShapes(WorksheetDrawing worksheetDrawing)
        {
            txtResults.AppendText("=== �}�`�\���̒��� ===\n");
            var childElements = worksheetDrawing.ChildElements.ToList();
            txtResults.AppendText($"WorksheetDrawing �̒����̎q�v�f��: {childElements.Count}\n");

            foreach (var element in childElements)
            {
                txtResults.AppendText($"�v�f�^�C�v: {element.GetType().Name}\n");
                if (element is TwoCellAnchor anchor)
                {
                    LogTwoCellAnchor(anchor);
                }
                // AbsoluteAnchor �� OneCellAnchor �����݂����邪�A����̗v���ł� TwoCellAnchor ���� GroupShape ����
            }
            txtResults.AppendText("=== �������� ===\n\n");
        }

        private void LogTwoCellAnchor(TwoCellAnchor anchor)
        {
            txtResults.AppendText("  TwoCellAnchor ���e:\n");
            foreach (var child in anchor.ChildElements)
            {
                txtResults.AppendText($"    {child.GetType().Name}\n");
                if (child is Shape shape) // ��O���[�v��Shape
                {
                    LogShapeDetails(shape, "      (�P��Shape) ");
                }
                else if (child is GroupShape groupShape)
                {
                    txtResults.AppendText("      GroupShape ���e:\n");
                    var shapesInGroup = groupShape.Elements<Shape>().ToList();
                    txtResults.AppendText($"        �O���[�v��Shape��: {shapesInGroup.Count}\n");
                    foreach (var groupedShape in shapesInGroup)
                    {
                        LogShapeDetails(groupedShape, "          (�O���[�v��Shape) ");
                    }
                }
                // ���ɂ� ConnectionShape �Ȃǂ����蓾��
            }
        }

        private void LogShapeDetails(Shape shape, string indent)
        {
            // Shape ID �� Name ������΃��O�ɏo�� (�f�o�b�O�p)
            var nonVisualSpPr = shape.NonVisualShapeProperties;
            if (nonVisualSpPr?.NonVisualDrawingProperties?.Id != null)
            {
                 txtResults.AppendText($"{indent}Shape ID: {nonVisualSpPr.NonVisualDrawingProperties.Id.Value}\n");
            }
            if (nonVisualSpPr?.NonVisualDrawingProperties?.Name != null)
            {
                 txtResults.AppendText($"{indent}Shape Name: {nonVisualSpPr.NonVisualDrawingProperties.Name.Value}\n");
            }

            var drawingTextBody = shape.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>();// Shape ������ TextBody (DocumentFormat.OpenXml.Drawing.TextBody)
            if (drawingTextBody != null)
            {
                var extractedText = ExtractTextFromDrawingTextBody(drawingTextBody);
                if (!string.IsNullOrWhiteSpace(extractedText))
                {
                    txtResults.AppendText($"{indent}�e�L�X�g: '{extractedText.Replace("\n", "\\n")}'\n");
                }
                else
                {
                    txtResults.AppendText($"{indent}�e�L�X�g�Ȃ� (Drawing.TextBody �͑���)\n");
                }
            }
            else
            {
                 txtResults.AppendText($"{indent}�e�L�X�g�Ȃ� (Drawing.TextBody �� null)\n");
            }
        }

        private TableInfo ExtractTableInfoFromGroup(GroupShape groupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromGroup �J�n\n");
            var shapesInGroup = groupShape.Elements<Shape>().ToList(); // Drawing.Shape
            txtResults.AppendText($"    �O���[�v����Shape��: {shapesInGroup.Count}\n");

            // �v��: �e�L�X�g�{�b�N�X1(�e�[�u����)�ƃe�L�X�g�{�b�N�X2(�J������)���O���[�v�� [cite: 1]
            if (shapesInGroup.Count != 2)
            {
                txtResults.AppendText($"    �O���[�v����Shape����2�ł͂���܂��� (���ۂ�{shapesInGroup.Count}��)�B�������X�L�b�v���܂��B\n");
                return null;
            }

            var textDataFromShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < shapesInGroup.Count; i++)
            {
                var shape = shapesInGroup[i];
                txtResults.AppendText($"    �O���[�v��Shape {i+1} �̃e�L�X�g���o���s...\n");
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
                        txtResults.AppendText($"      ���o�e�L�X�g: '{textContent.Replace("\n", "\\n")}' (�L���s��: {lines.Length})\n");
                    }
                    else
                    {
                        txtResults.AppendText($"      �e�L�X�g�͋�܂��͋󔒂ł��� (Drawing.TextBody �͑���)�B\n");
                    }
                }
                else
                {
                     txtResults.AppendText($"      Drawing.TextBody �� null �ł����B\n");
                }
                textDataFromShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            
            ShapeTextParseResult tableShapeData = null;
            ShapeTextParseResult columnsShapeData = null;

            // 1�ڂ�Shape���e�[�u����(1�s)�A2�ڂ�Shape���J������(1�s�ȏ�)�̃p�^�[��
            if (textDataFromShapes[0].Lines.Length == 1 && textDataFromShapes[1].Lines.Length >= 1)
            {
                tableShapeData = textDataFromShapes[0];
                columnsShapeData = textDataFromShapes[1];
                txtResults.AppendText("      �p�^�[�����v: Shape1=�e�[�u����, Shape2=�J������\n");
            }
            // 2�ڂ�Shape���e�[�u����(1�s)�A1�ڂ�Shape���J������(1�s�ȏ�)�̃p�^�[��
            else if (textDataFromShapes[1].Lines.Length == 1 && textDataFromShapes[0].Lines.Length >= 1)
            {
                tableShapeData = textDataFromShapes[1];
                columnsShapeData = textDataFromShapes[0];
                txtResults.AppendText("      �p�^�[�����v: Shape2=�e�[�u����, Shape1=�J������\n");
            }
            else
            {
                txtResults.AppendText($"    �e�[�u����(1�s)�ƃJ������(1�s�ȏ�)�̖��m�ȑg�ݍ��킹��������܂���B\n");
                txtResults.AppendText($"      Shape1 �L���s��: {textDataFromShapes[0].Lines.Length} (���e�L�X�g: '{textDataFromShapes[0].OriginalText.Replace("\n", "\\n")}')\n");
                txtResults.AppendText($"      Shape2 �L���s��: {textDataFromShapes[1].Lines.Length} (���e�L�X�g: '{textDataFromShapes[1].OriginalText.Replace("\n", "\\n")}')\n");
                return null;
            }

            string tableName = tableShapeData.Lines[0];
            List<string> columns = columnsShapeData.Lines.Distinct().ToList(); // �d���J�������͏���

            if (string.IsNullOrWhiteSpace(tableName))
            {
                txtResults.AppendText("    �e�[�u��������܂��͋󔒂̂��߁A�e�[�u�����Ƃ��ĕs�K�؂ł��B\n");
                return null;
            }
             if (!columns.Any())
            {
                txtResults.AppendText("    �J�������X�g����̂��߁A�e�[�u�����Ƃ��ĕs�K�؂ł��B\n");
                return null;
            }

            txtResults.AppendText($"    �e�[�u�����Ƃ��Ċm��: '{tableName}'\n");
            txtResults.AppendText($"    �J�������Ƃ��Ċm��: [{string.Join(", ", columns)}]\n");
            
            return new TableInfo
            {
                TableName = tableName,
                Columns = columns
            };
        }

        // DocumentFormat.OpenXml.Drawing.TextBody ����e�L�X�g�𒊏o
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
                 // ���s�R�[�h�� ParagraphProperties �� EndParagraphRunProperties �� <a:br/> �Ƃ��đ��݂���ꍇ�����邪�A
                 // �ʏ�AParagraph ���Ƃ����s�ɑΉ����邱�Ƃ������B
                 // �����ł́A�eParagraph�̓��e��P���ɘA�����A���Split����B
                textParts.Add(paragraphText);
            }
            // �eParagraph�͒ʏ�AExcel���ł̉��s�ɑΉ����邽�߁AJoin���ɉ��s������
            return string.Join("\n", textParts);
        }
        
        // �� ExtractTextFromTextBody (Spreadsheet.TextBody�p) �͍��񒼐ڎg��Ȃ����A�Q�l�Ƃ��Ďc���ꍇ�͕ʂ�
        // private string ExtractTextFromShape(Shape shape) �͕s�v�ɂȂ�A���� ExtractTextFromDrawingTextBody ���g��

        private string FormatResultsForDisplay(List<TableInfo> results) // txtResults �\���p
        {
            var output = new List<string>();
            output.Add("--- ���o�e�[�u���ꗗ ---");
            foreach (var table in results)
            {
                output.Add($"�e�[�u����: {table.TableName}");
                output.Add("  �J������:");
                foreach (var column in table.Columns)
                {
                    output.Add($"    - {column}");
                }
                output.Add(""); // ��s�ŋ�؂�
            }
            return string.Join("\n", output);
        }

        private void CreateOutputSheet(string filePath, string outputSheetName, List<TableInfo> results)
        {
            using (var document = SpreadsheetDocument.Open(filePath, true)) // �������݉\�ŊJ��
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null) throw new InvalidOperationException("WorkbookPart is null.");

                // �����̓����V�[�g���폜
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

                // �V�������[�N�V�[�g���쐬
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
                if (sheetData == null) // �ʏ�� AddNewPart �ō쐬�����͂�
                {
                     sheetData = newWorksheetPart.Worksheet.AppendChild(new SheetData());
                }

                // �w�b�_�[�s
                var headerRow = new Row() { RowIndex = 1U }; // 1-based index
                headerRow.Append(CreateCell("A", 1U, "#"));
                headerRow.Append(CreateCell("B", 1U, "num"));
                headerRow.Append(CreateCell("C", 1U, "�e�[�u����"));
                headerRow.Append(CreateCell("D", 1U, "�J������"));
                sheetData.Append(headerRow);

                // �f�[�^�s
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
                DataType = CellValues.String, // InlineString ���� String + SharedStringTable �����������ꍇ�����邪�A�P���ȃc�[���Ȃ� InlineString �ł���
                                            // DataType = CellValues.InlineString, // ������g���ꍇ�� InlineString �v�f���K�v
                CellValue = new CellValue(value)  // DataType = CellValues.String �̏ꍇ
                // InlineString = new InlineString(new Text(value)) // DataType = CellValues.InlineString �̏ꍇ
            };
        }
    }

    public class TableInfo
    {
        public string TableName { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
    }

    // �O���[�v���̊eShape����e�L�X�g���p�[�X�������ʂ�ێ���������N���X
    internal class ShapeTextParseResult
    {
        public string OriginalText { get; set; }
        public string[] Lines { get; set; } // �󔒍s�������ATrim���ꂽ�L���ȍs
    }
}