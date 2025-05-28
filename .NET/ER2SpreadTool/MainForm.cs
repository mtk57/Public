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
                
                if (results.Any())
                {
                    txtResults.AppendText($"\n���o����:\n{FormatResultsForDisplay(results)}");
                    CreateOutputSheet(txtFilePath.Text, "ER�}���o����", results); // Ensure this method works fine
                    lblStatus.Text = $"�������� - {results.Count}���̃e�[�u�����𒊏o���܂����B";
                    MessageBox.Show("�������������܂����B�V�K�V�[�g�uER�}���o���ʁv�Ɍ��ʂ��o�͂��܂����B", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    txtResults.AppendText("\n�����Ώۂ̐}�`��������܂���ł����B\n�f�o�b�O�����m�F���Ă��������B");
                    lblStatus.Text = "�����Ώۂ̐}�`��������܂���ł����B";
                    MessageBox.Show("�����Ώۂ̐}�`��������܂���ł����B���O���m�F���Ă��������B", "���", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "�����G���[";
                txtResults.AppendText($"\n�G���[���������܂���:\n{ex.ToString()}\n");
                MessageBox.Show($"�������ɃG���[���������܂���:\n{ex.Message}", "�G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    txtResults.AppendText("WorkbookPart ��������܂���B\n");
                    return results;
                }

                var sheet = workbookPart.Workbook.Descendants<Sheet>()
                    .FirstOrDefault(s => s.Name != null && s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                if (sheet == null) throw new ArgumentException($"�V�[�g '{sheetName}' ��������܂���B");

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
                LogAllShapes(worksheetDrawing);

                foreach (var anchor in worksheetDrawing.Elements<TwoCellAnchor>())
                {
                    // Iterate through DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape (xdr:grpSp) elements
                    foreach (var xdrGroupShape in anchor.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape>())
                    {
                        txtResults.AppendText("Spreadsheet.GroupShape (xdr:grpSp) ��TwoCellAnchor���Ŕ����B\n");
                        var tableInfo = ExtractTableInfoFromSpreadsheetGroup(xdrGroupShape);
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
                             txtResults.AppendText("  => Spreadsheet.GroupShape����e�[�u����񒊏o���s�B\n");
                        }
                    }
                }
            }
            return results;
        }
        
        private TableInfo ExtractTableInfoFromSpreadsheetGroup(DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromSpreadsheetGroup (xdr:grpSp) �J�n\n");

            // Scenario 1: The xdr:grpSp contains a nested a:grpSp (Drawing.GroupShape)
            var nestedDrawingGroup = xdrGroupShape.GetFirstChild<DocumentFormat.OpenXml.Drawing.GroupShape>();
            if (nestedDrawingGroup != null)
            {
                txtResults.AppendText("    �l�X�g���ꂽ Drawing.GroupShape (a:grpSp) �𔭌��B������������܂��B\n");
                return ExtractTableInfoFromDrawingGroupShape(nestedDrawingGroup);
            }

            // Scenario 2: The xdr:grpSp directly contains two xdr:sp (Spreadsheet.Shape) elements
            var xdrShapesInGroup = xdrGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().ToList();
            txtResults.AppendText($"    xdr:grpSp����Spreadsheet.Shape (xdr:sp) ��: {xdrShapesInGroup.Count}\n");

            if (xdrShapesInGroup.Count != 2)
            {
                txtResults.AppendText($"    xdr:grpSp����xdr:sp����2�ł͂���܂���B�����W�b�N�ł͏����s�B\n");
                return null;
            }

            var textDataFromXdrShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < xdrShapesInGroup.Count; i++)
            {
                var xdrShape = xdrShapesInGroup[i];
                txtResults.AppendText($"    xdr:sp {i + 1} �̃e�L�X�g���o���s...\n");
                string textContent = string.Empty;
                
                // Text in an xdr:sp is usually in its xdr:txBody (Spreadsheet.TextBody)
                var spreadsheetTextBody = xdrShape.TextBody; // This is DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody
                if (spreadsheetTextBody != null)
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) �𔭌��B\n");
                    textContent = ExtractTextFromSpreadsheetTextBody(spreadsheetTextBody);
                }
                else
                {
                    txtResults.AppendText("      Spreadsheet.TextBody (xdr:txBody) ��������܂���B\n");
                }

                string[] lines = new string[0];
                if (!string.IsNullOrWhiteSpace(textContent))
                {
                    lines = textContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                       .Select(l => l.Trim())
                                       .Where(l => !string.IsNullOrEmpty(l))
                                       .ToArray();
                    txtResults.AppendText($"      ���o�e�L�X�g: '{textContent.Replace("\n", "\\n")}' (�L���s��: {lines.Length})\n");
                }
                textDataFromXdrShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromXdrShapes, "xdr:sp�x�[�X");
        }

        private TableInfo ExtractTableInfoFromDrawingGroupShape(DocumentFormat.OpenXml.Drawing.GroupShape drawingGroupShape)
        {
            txtResults.AppendText("  ExtractTableInfoFromDrawingGroupShape (a:grpSp) �J�n\n");
            var shapesInDrawingGroup = drawingGroupShape.Elements<DocumentFormat.OpenXml.Drawing.Shape>().ToList(); // These are a:sp
            txtResults.AppendText($"    a:grpSp����Drawing.Shape (a:sp) ��: {shapesInDrawingGroup.Count}\n");

            if (shapesInDrawingGroup.Count != 2)
            {
                txtResults.AppendText($"    a:grpSp����a:sp����2�ł͂���܂���B\n");
                return null;
            }

            var textDataFromDrawingShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < shapesInDrawingGroup.Count; i++)
            {
                var drawingShape = shapesInDrawingGroup[i]; // This is a:sp
                txtResults.AppendText($"    a:sp {i + 1} �̃e�L�X�g���o���s...\n");
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
                        txtResults.AppendText($"      ���o�e�L�X�g: '{textContent.Replace("\n", "\\n")}' (�L���s��: {lines.Length})\n");
                    }
                }
                else
                {
                    txtResults.AppendText($"      a:sp {i + 1} �� Drawing.TextBody (a:txBody) ��������܂���B\n");
                }
                textDataFromDrawingShapes.Add(new ShapeTextParseResult { OriginalText = textContent, Lines = lines });
            }
            return CreateTableInfoFromTextData(textDataFromDrawingShapes, "a:sp�x�[�X");
        }
        
        private TableInfo CreateTableInfoFromTextData(List<ShapeTextParseResult> textDataList, string sourceDescription)
        {
            if (textDataList.Count != 2) {
                 txtResults.AppendText($"    {sourceDescription}: �e�L�X�g����2�擾�ł��܂���ł��� (���ۂ�{textDataList.Count}��)�B\n");
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
                txtResults.AppendText($"    �e�[�u����(1�s)�ƃJ������(1�s�ȏ�)�̖��m�ȑg�ݍ��킹��������܂��� ({sourceDescription})�B\n");
                txtResults.AppendText($"      Shape1 �L���s��: {textDataList[0].Lines.Length} (���e�L�X�g: '{textDataList[0].OriginalText.Replace("\n", "\\n")}')\n");
                txtResults.AppendText($"      Shape2 �L���s��: {textDataList[1].Lines.Length} (���e�L�X�g: '{textDataList[1].OriginalText.Replace("\n", "\\n")}')\n");
                return null;
            }

            // Ensure that the identified table name shape actually has content
            if (tableShapeData.Lines.Length == 0 || string.IsNullOrWhiteSpace(tableShapeData.Lines[0]))
            {
                txtResults.AppendText($"    {sourceDescription}: ���o���ꂽ�e�[�u��������ł��B\n");
                return null;
            }
            // Ensure that the identified columns shape actually has content
             if (columnsShapeData.Lines.Length == 0)
            {
                txtResults.AppendText($"    {sourceDescription}: ���o���ꂽ�J�������X�g����ł��B\n");
                return null;
            }


            string tableName = tableShapeData.Lines[0];
            List<string> columns = columnsShapeData.Lines.Distinct().ToList();
            
            txtResults.AppendText($"    {sourceDescription}: �e�[�u�����Ƃ��Ċm��: '{tableName}'\n");
            txtResults.AppendText($"    {sourceDescription}: �J�������Ƃ��Ċm��: [{string.Join(", ", columns)}]\n");

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
            txtResults.AppendText("=== �}�`�\���̒��� ===\n");
            var childElements = worksheetDrawing.ChildElements.ToList();
            txtResults.AppendText($"WorksheetDrawing �̒����̎q�v�f��: {childElements.Count}\n");

            foreach (var element in childElements)
            {
                txtResults.AppendText($"�v�f�^�C�v: {element.GetType().FullName}\n"); // Use FullName for clarity
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
                txtResults.AppendText($"    �q�v�f�^�C�v: {child.GetType().FullName}\n");
                if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape xdrShape)
                {
                    txtResults.AppendText("      �ڍ�: Spreadsheet.Shape (xdr:sp)\n");
                    var spTextBody = xdrShape.TextBody;
                    if(spTextBody != null) {
                        txtResults.AppendText($"        xdr:sp Text: '{ExtractTextFromSpreadsheetTextBody(spTextBody).Replace("\n", "\\n")}'\n");
                    }
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
                {
                    txtResults.AppendText("      �ڍ�: Spreadsheet.GroupShape (xdr:grpSp)\n");
                    foreach(var innerElement in xdrGroupShape.Elements<OpenXmlCompositeElement>()){
                        txtResults.AppendText($"        xdr:grpSp ���̎q�v�f: {innerElement.GetType().FullName}\n");
                        if(innerElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape innerXdrShape) {
                             var innerSpTextBody = innerXdrShape.TextBody;
                             if(innerSpTextBody != null) {
                                txtResults.AppendText($"          xdr:sp (��) Text: '{ExtractTextFromSpreadsheetTextBody(innerSpTextBody).Replace("\n", "\\n")}'\n");
                            }
                        } else if (innerElement is DocumentFormat.OpenXml.Drawing.GroupShape innerDrawingGroupShape) { // a:grpSp
                             txtResults.AppendText($"          Drawing.GroupShape (a:grpSp) (��)\n");
                        }
                    }
                }
            }
        }
        
        private string FormatResultsForDisplay(List<TableInfo> results)
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
                headerRow.Append(CreateCell("C", 1U, "�e�[�u����"));
                headerRow.Append(CreateCell("D", 1U, "�J������"));
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