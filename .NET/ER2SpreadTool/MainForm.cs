using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
// �� �ȉ���using�f�B���N�e�B�u��ǉ�
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace ER2SpreadTool
{
    // �ݒ��ێ�����N���X (DataContract�������g�p)
    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public string LastFilePath { get; set; }

        [DataMember]
        public string LastSheetName { get; set; }
    }

    public partial class MainForm : Form
    {
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();
            settingsFilePath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "er2spreadtool_settings.json");
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void LoadSettings()
        {
            if (!File.Exists(settingsFilePath))
            {
                txtResults.AppendText("�ݒ�t�@�C����������܂���ł����B����N���܂��͐ݒ�t�@�C�����폜����Ă��܂��B\n");
                return;
            }

            try
            {
                using (FileStream fs = new FileStream(settingsFilePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length == 0) // ��̃t�@�C���̏ꍇ
                    {
                        txtResults.AppendText("�ݒ�t�@�C���͋�ł��B\n");
                        return;
                    }
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    AppSettings settings = (AppSettings)serializer.ReadObject(fs);

                    if (settings != null)
                    {
                        txtFilePath.Text = settings.LastFilePath;
                        txtSheetName.Text = settings.LastSheetName;
                        txtResults.AppendText("�O��I�����̐ݒ��ǂݍ��݂܂����B\n");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�ݒ�t�@�C���̓ǂݍ��ݒ��ɃG���[���������܂���:\n{ex.Message}", "�ݒ�ǂݍ��݃G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                txtResults.AppendText($"�ݒ�t�@�C���̓ǂݍ��݃G���[: {ex.Message}\n");
            }
        }

        private void SaveSettings()
        {
            AppSettings settings = new AppSettings
            {
                LastFilePath = txtFilePath.Text,
                LastSheetName = txtSheetName.Text
            };

            try
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    serializer.WriteObject(ms, settings);
                    byte[] jsonBytes = ms.ToArray();
                    // �C���f���g���ꂽJSON�Ƃ��ď����o�����߂ɁA��x������ɕϊ����Ď蓮�ŃC���f���g���ɂ��邩�A
                    // �܂���DataContractJsonSerializerSettings�ŃC���f���g�I�v�V������T���i�W���ł͒P���ȃC���f���g�͓���j
                    // �����ł͒P���ȏ����o���Ƃ��܂��B
                    File.WriteAllBytes(settingsFilePath, jsonBytes);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"�ݒ�t�@�C���̕ۑ����ɃG���[���������܂���:\n{ex.Message}", "�ݒ�ۑ��G���[", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // �K�v�ł���΃��O�ɏ����o���Ȃǂ̏�����ǉ�
            }
        }

        // btnBrowse_Click ���\�b�h�������ύX���āA�ǂݍ��񂾃p�X�𗘗p
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel�t�@�C�� (*.xlsx)|*.xlsx";
                openFileDialog.Title = "�����Ώۂ�Excel�t�@�C����I�����Ă�������";

                // �O��l������΂���������l�Ƃ���
                if (!string.IsNullOrWhiteSpace(txtFilePath.Text))
                {
                    if (File.Exists(txtFilePath.Text))
                    {
                        openFileDialog.InitialDirectory = Path.GetDirectoryName(txtFilePath.Text);
                        openFileDialog.FileName = Path.GetFileName(txtFilePath.Text);
                    }
                    else if (Directory.Exists(txtFilePath.Text)) // �����p�X���f�B���N�g���Ȃ�
                    {
                        openFileDialog.InitialDirectory = txtFilePath.Text;
                    }
                }


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
                // txtResults.Clear(); // ���O�̃N���A�^�C�~���O�͗v�]�ɉ����Ē���
                Application.DoEvents();

                var results = ProcessExcelFile(txtFilePath.Text, txtSheetName.Text);

                if (results.Any())
                {
                    txtResults.AppendText($"\n���o����:\n{FormatResultsForDisplay(results)}");
                    
                    // TSV�t�@�C���o�́i���̓t�@�C����ύX���Ȃ��j
                    string outputFilePath = CreateTsvOutput(txtFilePath.Text, results);
                    
                    lblStatus.Text = $"�������� - {results.Count}���̃e�[�u�����𒊏o���܂����B";
                    MessageBox.Show($"�������������܂����B\nTSV�t�@�C�����o�͂��܂���:\n{outputFilePath}", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            if (xdrShapesInGroup.Count < 2) // Needs at least one for table name, one for columns
            {
                txtResults.AppendText($"    xdr:grpSp���ɏ��Ȃ��Ƃ�2��xdr:sp���K�v�ł� (�e�[�u�����p1�A�J�������p1�ȏ�)�B���ۂ�{xdrShapesInGroup.Count}�B\n");
                return null;
            }

            var textDataFromXdrShapes = new List<ShapeTextParseResult>();
            for (int i = 0; i < xdrShapesInGroup.Count; i++)
            {
                var xdrShape = xdrShapesInGroup[i];
                txtResults.AppendText($"    xdr:sp {i + 1} �̃e�L�X�g���o���s...\n");
                string textContent = string.Empty;

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

            if (shapesInDrawingGroup.Count < 2) // Needs at least one for table name, one for columns
            {
                txtResults.AppendText($"    a:grpSp���ɏ��Ȃ��Ƃ�2��a:sp���K�v�ł� (�e�[�u�����p1�A�J�������p1�ȏ�)�B���ۂ�{shapesInDrawingGroup.Count}�B\n");
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
            txtResults.AppendText($"    {sourceDescription}: CreateTableInfoFromTextData �J�n�BShapeTextParseResult ��: {textDataList.Count}\n");

            if (textDataList.Count < 2)
            {
                txtResults.AppendText($"    {sourceDescription}: ���Ȃ��Ƃ�2�̃e�L�X�g��񂪕K�v�ł� (�e�[�u�����p1�A�J�������p1�ȏ�)�B���ۂ�{textDataList.Count}�B\n");
                return null;
            }

            // Identify potential table name shapes (those with exactly one line of non-empty text)
            List<ShapeTextParseResult> potentialTableShapes = textDataList
                .Where(s => s.Lines != null && s.Lines.Length == 1 && !string.IsNullOrWhiteSpace(s.Lines[0]))
                .ToList();

            ShapeTextParseResult tableShapeData = null;
            List<string> columnLinesAggregated = new List<string>();

            if (potentialTableShapes.Count == 1)
            {
                tableShapeData = potentialTableShapes[0];
                txtResults.AppendText($"    {sourceDescription}: �e�[�u�������̃V�F�C�v�𔭌� (���e�L�X�g: '{tableShapeData.OriginalText.Replace("\n", "\\n")}')\n");

                // All other shapes contribute to columns
                foreach (var shapeData in textDataList)
                {
                    if (shapeData != tableShapeData) // If it's not the table name shape
                    {
                        if (shapeData.Lines != null && shapeData.Lines.Length > 0)
                        {
                            columnLinesAggregated.AddRange(shapeData.Lines);
                            txtResults.AppendText($"    {sourceDescription}: �J�������̃V�F�C�v����s��ǉ� (���e�L�X�g: '{shapeData.OriginalText.Replace("\n", "\\n")}', �ǉ��s��: {shapeData.Lines.Length})\n");
                        }
                        else
                        {
                            txtResults.AppendText($"    {sourceDescription}: �J�������V�F�C�v�ɗL���ȍs�Ȃ��A�܂��͋�e�L�X�g (���e�L�X�g: '{shapeData.OriginalText.Replace("\n", "\\n")}')\n");
                        }
                    }
                }
            }
            else
            {
                txtResults.AppendText($"    {sourceDescription}: �e�[�u�����ƂȂ�V�F�C�v (�e�L�X�g1�s�̂�) ��1�ł���K�v������܂����A{potentialTableShapes.Count}������܂����B\n");
                txtResults.AppendText($"      �����ΏۃV�F�C�v��: {textDataList.Count}\n");
                for (int i = 0; i < textDataList.Count; i++)
                {
                    var currentShape = textDataList[i];
                    txtResults.AppendText($"        Shape {i + 1} �L���s��: {(currentShape.Lines?.Length ?? 0)} (���e�L�X�g: '{currentShape.OriginalText.Replace("\n", "\\n")}')\n");
                }
                return null;
            }

            // Validate that a table name was indeed found and is not empty.
            if (tableShapeData == null) // Should ideally not happen if potentialTableShapes.Count == 1
            {
                txtResults.AppendText($"    {sourceDescription}: �e�[�u�����V�F�C�v�̓���Ɏ��s���܂��� (�����G���[)�B\n");
                return null;
            }
            string tableName = tableShapeData.Lines[0].Trim();


            // Validate that column data was found.
            if (columnLinesAggregated.Count == 0)
            {
                txtResults.AppendText($"    {sourceDescription}: ���o���ꂽ�J�������X�g����ł� (�e�[�u����: '{tableName}')�B\n");
                return null;
            }

            // Process and deduplicate column names
            List<string> finalColumns = columnLinesAggregated
                                        .Where(l => !string.IsNullOrWhiteSpace(l)) // Ensure no empty/whitespace lines become columns
                                        .Select(l => l.Trim())                     // Trim whitespace from each column name
                                        .Distinct(StringComparer.OrdinalIgnoreCase) // Deduplicate column names (case-insensitive)
                                        .ToList();

            if (finalColumns.Count == 0) // After trimming and distinct, it might become empty
            {
                txtResults.AppendText($"    {sourceDescription}: �L���ȃJ�����������o�E������A��ɂȂ�܂��� (�e�[�u����: '{tableName}')�B\n");
                return null;
            }

            txtResults.AppendText($"    {sourceDescription}: �e�[�u�����Ƃ��Ċm��: '{tableName}'\n");
            txtResults.AppendText($"    {sourceDescription}: �J�������Ƃ��Ċm�� ({finalColumns.Count}��): [{string.Join(", ", finalColumns)}]\n");

            return new TableInfo { TableName = tableName, Columns = finalColumns };
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
                    if (spTextBody != null)
                    {
                        txtResults.AppendText($"        xdr:sp Text: '{ExtractTextFromSpreadsheetTextBody(spTextBody).Replace("\n", "\\n")}'\n");
                    }
                }
                else if (child is DocumentFormat.OpenXml.Drawing.Spreadsheet.GroupShape xdrGroupShape)
                {
                    txtResults.AppendText("      �ڍ�: Spreadsheet.GroupShape (xdr:grpSp)\n");
                    foreach (var innerElement in xdrGroupShape.Elements<OpenXmlCompositeElement>())
                    {
                        txtResults.AppendText($"        xdr:grpSp ���̎q�v�f: {innerElement.GetType().FullName}\n");
                        if (innerElement is DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape innerXdrShape)
                        {
                            var innerSpTextBody = innerXdrShape.TextBody;
                            if (innerSpTextBody != null)
                            {
                                txtResults.AppendText($"          xdr:sp (��) Text: '{ExtractTextFromSpreadsheetTextBody(innerSpTextBody).Replace("\n", "\\n")}'\n");
                            }
                        }
                        else if (innerElement is DocumentFormat.OpenXml.Drawing.GroupShape innerDrawingGroupShape)
                        { // a:grpSp
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

        // TSV�t�@�C���o�̓��\�b�h�i�V�K�ǉ��j
        private string CreateTsvOutput(string inputFilePath, List<TableInfo> results)
        {
            // ���̓t�@�C���Ɠ����t�H���_��TSV�t�@�C�����쐬
            string inputDirectory = Path.GetDirectoryName(inputFilePath);
            
            // ���ݓ�����YYYYMMDD_hhmmssfff�`���Ńt�H�[�}�b�g
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
            string outputFileName = $"{timestamp}.tsv";
            string outputFilePath = Path.Combine(inputDirectory, outputFileName);

            try
            {
                using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    // �w�b�_�[�s���o��
                    writer.WriteLine("#\tnum\t�e�[�u����\t�J������");

                    // �f�[�^�s���o��
                    int overallCounter = 1;
                    foreach (var table in results)
                    {
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            writer.WriteLine($"{overallCounter}\t{i + 1}\t{table.TableName}\t{table.Columns[i]}");
                            overallCounter++;
                        }
                    }
                }

                txtResults.AppendText($"TSV�t�@�C�����o�͂��܂���: {outputFilePath}\n");
                return outputFilePath;
            }
            catch (Exception ex)
            {
                string errorMessage = $"TSV�t�@�C���̏o�͒��ɃG���[���������܂���: {ex.Message}";
                txtResults.AppendText($"{errorMessage}\n");
                MessageBox.Show(errorMessage, "�o�̓G���[", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
    }

    // TableInfo �� ShapeTextParseResult �͊����̂܂�
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