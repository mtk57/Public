using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Excel2Tsv
{
    public partial class Form1 : Form
    {
        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx",
            ".xlsm",
            ".xlsb"
        };

        public Form1()
        {
            InitializeComponent();
            InitializeUi();
        }

        private void InitializeUi()
        {
            label1.Text = "Excel ファイルパス (xlsx/xlsm/xlsb) ※必須";

            txtExcelFilePath.AllowDrop = true;
            txtTsvDirPath.AllowDrop = true;

            txtExcelFilePath.DragEnter += FileDropTarget_DragEnter;
            txtExcelFilePath.DragDrop += TxtExcelFilePath_DragDrop;
            txtTsvDirPath.DragEnter += FileDropTarget_DragEnter;
            txtTsvDirPath.DragDrop += TxtTsvDirPath_DragDrop;

            btnRefExcelFilePath.Click += BtnRefExcelFilePath_Click;
            btnRefTsvDirPath.Click += BtnRefTsvDirPath_Click;
            btnStart.Click += BtnStart_Click;

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { Name = "SheetName", HeaderText = "シート名", SortMode = DataGridViewColumnSortMode.NotSortable });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { Name = "TsvFileName", HeaderText = "TSVファイル名", SortMode = DataGridViewColumnSortMode.NotSortable });
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = true;
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridView1.KeyDown += DataGridView1_KeyDown;
        }

        private void FileDropTarget_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
                return;
            }

            e.Effect = DragDropEffects.None;
        }

        private void TxtExcelFilePath_DragDrop(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0)
            {
                return;
            }

            var excelPath = files.FirstOrDefault(IsSupportedExcelFile);
            if (excelPath == null)
            {
                MessageBox.Show("xlsx / xlsm / xlsb のファイルをドロップしてください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtExcelFilePath.Text = excelPath;
        }

        private void TxtTsvDirPath_DragDrop(object sender, DragEventArgs e)
        {
            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0)
            {
                return;
            }

            var directory = files.FirstOrDefault(Directory.Exists);
            if (directory == null)
            {
                MessageBox.Show("フォルダをドロップしてください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtTsvDirPath.Text = directory;
        }

        private void BtnRefExcelFilePath_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel ファイル (*.xlsx;*.xlsm;*.xlsb)|*.xlsx;*.xlsm;*.xlsb";
                dialog.CheckFileExists = true;
                dialog.Multiselect = false;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtExcelFilePath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRefTsvDirPath_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "TSV 出力フォルダを選択してください";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtTsvDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void DataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                PasteClipboardToGrid();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.C)
            {
                var data = dataGridView1.GetClipboardContent();
                if (data != null)
                {
                    Clipboard.SetDataObject(data);
                }

                e.Handled = true;
            }
        }

        private void PasteClipboardToGrid()
        {
            if (!Clipboard.ContainsText())
            {
                return;
            }

            var text = Clipboard.GetText();
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var normalized = text.Replace("\r\n", "\n");
            var lines = normalized.Split('\n');
            if (lines.Length > 0 && lines[lines.Length - 1].Length == 0)
            {
                Array.Resize(ref lines, lines.Length - 1);
            }

            var startRow = dataGridView1.CurrentCell != null ? dataGridView1.CurrentCell.RowIndex : 0;
            var startColumn = dataGridView1.CurrentCell != null ? dataGridView1.CurrentCell.ColumnIndex : 0;

            for (var rowOffset = 0; rowOffset < lines.Length; rowOffset++)
            {
                var targetRowIndex = startRow + rowOffset;
                while (targetRowIndex >= dataGridView1.Rows.Count - 1)
                {
                    dataGridView1.Rows.Add();
                }

                var values = lines[rowOffset].Split('\t');
                for (var colOffset = 0; colOffset < values.Length; colOffset++)
                {
                    var targetColIndex = startColumn + colOffset;
                    if (targetColIndex >= dataGridView1.Columns.Count)
                    {
                        break;
                    }

                    dataGridView1.Rows[targetRowIndex].Cells[targetColIndex].Value = values[colOffset];
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                var outputDirectory = ConvertExcelSheetsToTsv();
                MessageBox.Show("完了しました。\r\n出力先: " + outputDirectory, "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnStart.Enabled = true;
            }
        }

        private string ConvertExcelSheetsToTsv()
        {
            var excelFilePath = (txtExcelFilePath.Text ?? string.Empty).Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                throw new InvalidOperationException("Excel ファイルパスを入力してください。");
            }

            if (!File.Exists(excelFilePath))
            {
                throw new InvalidOperationException("Excel ファイルが存在しません: " + excelFilePath);
            }

            if (!IsSupportedExcelFile(excelFilePath))
            {
                throw new InvalidOperationException("サポート対象の拡張子は .xlsx / .xlsm / .xlsb のみです。");
            }

            var mappings = GetSheetMappings();
            if (mappings.Count == 0)
            {
                throw new InvalidOperationException("dataGridView1 に1行以上の変換設定を入力してください。");
            }

            var outputDirectory = ResolveOutputDirectory(excelFilePath, txtTsvDirPath.Text);

            using (var document = SpreadsheetDocument.Open(excelFilePath, false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null)
                {
                    throw new InvalidOperationException("Excel ファイルの読み込みに失敗しました。");
                }

                var sharedStringItems = workbookPart.SharedStringTablePart != null
                    && workbookPart.SharedStringTablePart.SharedStringTable != null
                    ? workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ToList()
                    : new List<SharedStringItem>();

                foreach (var mapping in mappings)
                {
                    var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                        .FirstOrDefault(x => string.Equals(x.Name != null ? x.Name.Value : string.Empty, mapping.SheetName, StringComparison.OrdinalIgnoreCase));

                    if (sheet == null)
                    {
                        throw new InvalidOperationException("シートが見つかりません: " + mapping.SheetName);
                    }

                    if (sheet.Id == null)
                    {
                        throw new InvalidOperationException("シートIDが不正です: " + mapping.SheetName);
                    }

                    var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    if (worksheetPart == null)
                    {
                        throw new InvalidOperationException("シートを開けませんでした: " + mapping.SheetName);
                    }

                    var outputFilePath = Path.Combine(outputDirectory, NormalizeTsvFileName(mapping.TsvFileName));
                    WriteWorksheetAsTsv(worksheetPart, sharedStringItems, outputFilePath);
                }
            }

            return outputDirectory;
        }

        private string ResolveOutputDirectory(string excelFilePath, string tsvDirectoryPath)
        {
            var candidate = (tsvDirectoryPath ?? string.Empty).Trim().Trim('"');
            if (!string.IsNullOrWhiteSpace(candidate) && Directory.Exists(candidate))
            {
                return candidate;
            }

            var excelDirectory = Path.GetDirectoryName(excelFilePath);
            if (string.IsNullOrEmpty(excelDirectory))
            {
                throw new InvalidOperationException("Excel ファイルの親フォルダを取得できません。");
            }

            var fallbackDirectory = Path.Combine(excelDirectory, Guid.NewGuid().ToString());
            Directory.CreateDirectory(fallbackDirectory);
            txtTsvDirPath.Text = fallbackDirectory;
            return fallbackDirectory;
        }

        private List<SheetMapping> GetSheetMappings()
        {
            var mappings = new List<SheetMapping>();
            for (var i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var row = dataGridView1.Rows[i];
                if (row.IsNewRow)
                {
                    continue;
                }

                var sheetName = Convert.ToString(row.Cells[0].Value) ?? string.Empty;
                var tsvFileName = Convert.ToString(row.Cells[1].Value) ?? string.Empty;
                sheetName = sheetName.Trim();
                tsvFileName = tsvFileName.Trim();

                if (sheetName.Length == 0 && tsvFileName.Length == 0)
                {
                    continue;
                }

                if (sheetName.Length == 0)
                {
                    throw new InvalidOperationException((i + 1) + "行目: シート名を入力してください。");
                }

                if (tsvFileName.Length == 0)
                {
                    throw new InvalidOperationException((i + 1) + "行目: TSVファイル名を入力してください。");
                }

                mappings.Add(new SheetMapping
                {
                    SheetName = sheetName,
                    TsvFileName = tsvFileName
                });
            }

            return mappings;
        }

        private static string NormalizeTsvFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new InvalidOperationException("TSVファイル名が空です。");
            }

            var trimmed = fileName.Trim();
            if (Path.GetFileName(trimmed) != trimmed)
            {
                throw new InvalidOperationException("TSVファイル名にフォルダは指定できません: " + fileName);
            }

            if (trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new InvalidOperationException("TSVファイル名に使用できない文字が含まれています: " + fileName);
            }

            if (string.IsNullOrEmpty(Path.GetExtension(trimmed)))
            {
                return trimmed + ".tsv";
            }

            return trimmed;
        }

        private static bool IsSupportedExcelFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            var extension = Path.GetExtension(path);
            return SupportedExtensions.Contains(extension);
        }

        private static void WriteWorksheetAsTsv(WorksheetPart worksheetPart, IReadOnlyList<SharedStringItem> sharedStringItems, string outputFilePath)
        {
            var sheetData = worksheetPart.Worksheet != null ? worksheetPart.Worksheet.GetFirstChild<SheetData>() : null;

            using (var writer = new StreamWriter(outputFilePath, false, new UTF8Encoding(false)))
            {
                if (sheetData == null)
                {
                    return;
                }

                uint expectedRowIndex = 1;
                foreach (var row in sheetData.Elements<Row>())
                {
                    var rowIndex = row.RowIndex != null ? row.RowIndex.Value : expectedRowIndex;
                    while (expectedRowIndex < rowIndex)
                    {
                        writer.Write("\r\n");
                        expectedRowIndex++;
                    }

                    var rowValues = ReadRowValues(sharedStringItems, row);
                    writer.Write(string.Join("\t", rowValues.Select(EscapeTsvValue)));
                    writer.Write("\r\n");
                    expectedRowIndex = rowIndex + 1;
                }
            }
        }

        private static List<string> ReadRowValues(IReadOnlyList<SharedStringItem> sharedStringItems, Row row)
        {
            var values = new List<string>();
            var currentColumnIndex = 1;

            foreach (var cell in row.Elements<Cell>())
            {
                var columnIndex = GetColumnIndexFromCellReference(cell.CellReference);
                while (currentColumnIndex < columnIndex)
                {
                    values.Add(string.Empty);
                    currentColumnIndex++;
                }

                values.Add(GetCellValue(sharedStringItems, cell));
                currentColumnIndex = columnIndex + 1;
            }

            return values;
        }

        private static int GetColumnIndexFromCellReference(StringValue cellReference)
        {
            if (cellReference == null || string.IsNullOrEmpty(cellReference.Value))
            {
                return 1;
            }

            var columnName = new string(cellReference.Value.TakeWhile(char.IsLetter).ToArray());
            if (string.IsNullOrEmpty(columnName))
            {
                return 1;
            }

            var sum = 0;
            foreach (var character in columnName.ToUpperInvariant())
            {
                sum *= 26;
                sum += character - 'A' + 1;
            }

            return sum;
        }

        private static string GetCellValue(IReadOnlyList<SharedStringItem> sharedStringItems, Cell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            var rawValue = cell.CellValue != null ? cell.CellValue.Text : string.Empty;
            if (cell.DataType == null)
            {
                return rawValue;
            }

            var dataType = cell.DataType.Value;

            if (dataType == CellValues.SharedString)
            {
                return GetSharedStringValue(sharedStringItems, rawValue);
            }

            if (dataType == CellValues.Boolean)
            {
                return rawValue == "1" ? "TRUE" : "FALSE";
            }

            if (dataType == CellValues.InlineString)
            {
                return cell.InnerText ?? string.Empty;
            }

            return rawValue;
        }

        private static string GetSharedStringValue(IReadOnlyList<SharedStringItem> sharedStringItems, string indexText)
        {
            int index;
            if (!int.TryParse(indexText, out index))
            {
                return string.Empty;
            }

            if (index < 0 || index >= sharedStringItems.Count)
            {
                return string.Empty;
            }

            return sharedStringItems[index].InnerText ?? string.Empty;
        }

        private static string EscapeTsvValue(string value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            if (value.IndexOfAny(new[] { '\t', '\r', '\n', '"' }) < 0)
            {
                return value;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private sealed class SheetMapping
        {
            public string SheetName { get; set; }

            public string TsvFileName { get; set; }
        }

        private void Form1_Load ( object sender, EventArgs e )
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
        }
    }
}
