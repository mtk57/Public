using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel2Tsv
{
    public partial class Form1 : Form
    {
        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx",
            ".xlsm"
        };

        private const int ProgressScale = 1000;
        private const int ProgressReportRowInterval = 2000;
        private const int MaxParallelSheets = 4;

        private CancellationTokenSource _abortTokenSource;
        private bool _isRunning;

        public Form1()
        {
            InitializeComponent();
            InitializeUi();
        }

        private void InitializeUi()
        {
            label1.Text = "Excel ファイルパス (xlsx/xlsm) ※必須";

            txtExcelFilePath.AllowDrop = true;
            txtTsvDirPath.AllowDrop = true;

            txtExcelFilePath.DragEnter += FileDropTarget_DragEnter;
            txtExcelFilePath.DragDrop += TxtExcelFilePath_DragDrop;
            txtTsvDirPath.DragEnter += FileDropTarget_DragEnter;
            txtTsvDirPath.DragDrop += TxtTsvDirPath_DragDrop;

            btnRefExcelFilePath.Click += BtnRefExcelFilePath_Click;
            btnRefTsvDirPath.Click += BtnRefTsvDirPath_Click;
            btnStart.Click += BtnStart_Click;
            btnAbort.Click += BtnAbort_Click;

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { Name = "SheetName", HeaderText = "シート名", SortMode = DataGridViewColumnSortMode.NotSortable });
            dataGridView1.Columns.Add(new DataGridViewTextBoxColumn { Name = "TsvFileName", HeaderText = "TSVファイル名", SortMode = DataGridViewColumnSortMode.NotSortable });
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.MultiSelect = true;
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridView1.KeyDown += DataGridView1_KeyDown;

            progressBar1.Minimum = 0;
            progressBar1.Maximum = ProgressScale;
            progressBar1.Value = 0;
            btnAbort.Enabled = false;
            UpdateProgressUi(new ConversionProgress { TotalSheets = 0, CompletedSheets = 0, TotalEstimatedRows = 0, ProcessedRows = 0, Message = string.Empty });
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
                MessageBox.Show("xlsx / xlsm のファイルをドロップしてください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                dialog.Filter = "Excel ファイル (*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
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
            if (e.KeyCode == Keys.Delete)
            {
                ClearSelectedCells();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

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

        private async void BtnStart_Click(object sender, EventArgs e)
        {
            if (_isRunning)
            {
                return;
            }

            ConversionRequest request;
            try
            {
                request = BuildConversionRequest();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _isRunning = true;
            _abortTokenSource = new CancellationTokenSource();
            SetRunningState(true);
            UpdateProgressUi(new ConversionProgress
            {
                TotalSheets = request.Mappings.Count,
                CompletedSheets = 0,
                TotalEstimatedRows = Math.Max(1, request.Mappings.Count),
                ProcessedRows = 0,
                Message = "準備中"
            });

            var progress = new Progress<ConversionProgress>(UpdateProgressUi);
            try
            {
                var outputDirectory = await Task.Run(
                    () => ConvertExcelSheetsToTsv(request, _abortTokenSource.Token, progress),
                    _abortTokenSource.Token);

                MessageBox.Show("完了しました。\r\n出力先: " + outputDirectory, "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("処理を中止しました。", "中止", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (_abortTokenSource != null)
                {
                    _abortTokenSource.Dispose();
                    _abortTokenSource = null;
                }

                _isRunning = false;
                SetRunningState(false);
            }
        }

        private void BtnAbort_Click(object sender, EventArgs e)
        {
            if (_abortTokenSource == null || _abortTokenSource.IsCancellationRequested)
            {
                return;
            }

            btnAbort.Enabled = false;
            lblProgress.Text = "中止要求中...";
            _abortTokenSource.Cancel();
        }

        private void SetRunningState(bool running)
        {
            btnStart.Enabled = !running;
            btnAbort.Enabled = running;
            btnRefExcelFilePath.Enabled = !running;
            btnRefTsvDirPath.Enabled = !running;
            txtExcelFilePath.Enabled = !running;
            txtTsvDirPath.Enabled = !running;
            dataGridView1.Enabled = !running;
            Cursor = running ? Cursors.WaitCursor : Cursors.Default;
        }

        private void UpdateProgressUi(ConversionProgress progress)
        {
            var totalSheets = Math.Max(1, progress.TotalSheets);
            var completedSheets = Math.Max(0, Math.Min(progress.CompletedSheets, totalSheets));

            var totalRows = Math.Max(1L, progress.TotalEstimatedRows);
            var processedRows = Math.Max(0L, Math.Min(progress.ProcessedRows, totalRows));

            var ratio = (double)processedRows / totalRows;
            if (progress.TotalEstimatedRows <= 0)
            {
                ratio = (double)completedSheets / totalSheets;
            }

            var progressValue = (int)Math.Round(ratio * ProgressScale);
            progressBar1.Value = Math.Max(progressBar1.Minimum, Math.Min(progressBar1.Maximum, progressValue));

            var percent = ratio * 100.0;
            lblProgress.Text = string.Format(
                "{0}/{1} シート  {2:N0}/{3:N0} 行  {4:0.0}%",
                completedSheets,
                progress.TotalSheets,
                processedRows,
                progress.TotalEstimatedRows,
                percent);

            if (!string.IsNullOrWhiteSpace(progress.Message))
            {
                lblProgress.Text += "  " + progress.Message;
            }
        }

        private ConversionRequest BuildConversionRequest()
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
                throw new InvalidOperationException("サポート対象の拡張子は .xlsx / .xlsm のみです。");
            }

            var mappings = GetSheetMappings();
            if (mappings.Count == 0)
            {
                throw new InvalidOperationException("dataGridView1 に1行以上の変換設定を入力してください。");
            }

            var outputDirectory = ResolveOutputDirectory(excelFilePath, txtTsvDirPath.Text);
            txtTsvDirPath.Text = outputDirectory;

            return new ConversionRequest
            {
                ExcelFilePath = excelFilePath,
                OutputDirectory = outputDirectory,
                Mappings = mappings
            };
        }

        private static string ConvertExcelSheetsToTsv(ConversionRequest request, CancellationToken token, IProgress<ConversionProgress> progress)
        {
            token.ThrowIfCancellationRequested();

            var plan = BuildConversionPlan(request, token);
            var tracker = new ProgressTracker(plan.TotalEstimatedRows, plan.WorkItems.Count, progress);
            tracker.Report("開始");

            var options = new ParallelOptions
            {
                CancellationToken = token,
                MaxDegreeOfParallelism = GetMaxDegreeOfParallelism(plan.WorkItems.Count)
            };

            try
            {
                Parallel.ForEach(plan.WorkItems, options, workItem =>
                {
                    ConvertSingleSheet(
                        request.ExcelFilePath,
                        workItem,
                        plan.SharedStringValues,
                        token,
                        tracker);
                });
            }
            catch (AggregateException ex)
            {
                var flattened = ex.Flatten();
                var nonCancel = flattened.InnerExceptions.FirstOrDefault(x => !(x is OperationCanceledException));
                if (nonCancel != null)
                {
                    throw nonCancel;
                }

                throw new OperationCanceledException(token);
            }

            tracker.Report("完了");
            return request.OutputDirectory;
        }

        private static ConversionPlan BuildConversionPlan(ConversionRequest request, CancellationToken token)
        {
            var workItems = new List<SheetWorkItem>(request.Mappings.Count);
            var outputPathSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var totalEstimatedRows = 0L;
            List<string> sharedStringValues;

            using (var document = SpreadsheetDocument.Open(request.ExcelFilePath, false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.Sheets == null)
                {
                    throw new InvalidOperationException("Excel ファイルの読み込みに失敗しました。");
                }

                sharedStringValues = LoadSharedStringValues(workbookPart);

                foreach (var mapping in request.Mappings)
                {
                    token.ThrowIfCancellationRequested();

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

                    var outputPath = Path.Combine(request.OutputDirectory, NormalizeTsvFileName(mapping.TsvFileName));
                    if (!outputPathSet.Add(outputPath))
                    {
                        throw new InvalidOperationException("TSVファイル名が重複しています: " + mapping.TsvFileName);
                    }

                    var estimatedRows = EstimateTotalRows(worksheetPart);
                    totalEstimatedRows += estimatedRows;

                    workItems.Add(new SheetWorkItem
                    {
                        SheetName = mapping.SheetName,
                        RelationshipId = sheet.Id.Value,
                        OutputFilePath = outputPath,
                        EstimatedRows = estimatedRows
                    });
                }
            }

            return new ConversionPlan
            {
                SharedStringValues = sharedStringValues,
                WorkItems = workItems,
                TotalEstimatedRows = Math.Max(1, totalEstimatedRows)
            };
        }

        private static int GetMaxDegreeOfParallelism(int workItemCount)
        {
            var cpuBound = Math.Max(1, Environment.ProcessorCount);
            var limit = Math.Min(MaxParallelSheets, cpuBound);
            return Math.Max(1, Math.Min(limit, workItemCount));
        }

        private static List<string> LoadSharedStringValues(WorkbookPart workbookPart)
        {
            if (workbookPart.SharedStringTablePart == null || workbookPart.SharedStringTablePart.SharedStringTable == null)
            {
                return new List<string>();
            }

            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                .Select(x => x.InnerText ?? string.Empty)
                .ToList();
        }

        private static long EstimateTotalRows(WorksheetPart worksheetPart)
        {
            var reference = worksheetPart.Worksheet != null
                && worksheetPart.Worksheet.SheetDimension != null
                && worksheetPart.Worksheet.SheetDimension.Reference != null
                ? worksheetPart.Worksheet.SheetDimension.Reference.Value
                : string.Empty;

            if (string.IsNullOrWhiteSpace(reference))
            {
                return 1;
            }

            var lastCellReference = reference;
            var separatorIndex = reference.LastIndexOf(':');
            if (separatorIndex >= 0 && separatorIndex < reference.Length - 1)
            {
                lastCellReference = reference.Substring(separatorIndex + 1);
            }

            var rowText = new string(lastCellReference.Where(char.IsDigit).ToArray());
            long rowNumber;
            if (long.TryParse(rowText, out rowNumber) && rowNumber > 0)
            {
                return rowNumber;
            }

            return 1;
        }

        private static void ConvertSingleSheet(
            string excelFilePath,
            SheetWorkItem workItem,
            IReadOnlyList<string> sharedStringValues,
            CancellationToken token,
            ProgressTracker tracker)
        {
            token.ThrowIfCancellationRequested();

            using (var document = SpreadsheetDocument.Open(excelFilePath, false))
            {
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    throw new InvalidOperationException("Excel ファイルの読み込みに失敗しました。");
                }

                var worksheetPart = workbookPart.GetPartById(workItem.RelationshipId) as WorksheetPart;
                if (worksheetPart == null)
                {
                    throw new InvalidOperationException("シートを開けませんでした: " + workItem.SheetName);
                }

                var pendingRows = 0L;
                WriteWorksheetAsTsv(
                    worksheetPart,
                    sharedStringValues,
                    workItem.OutputFilePath,
                    token,
                    rows =>
                    {
                        pendingRows += rows;
                        if (pendingRows >= ProgressReportRowInterval)
                        {
                            tracker.AddRows(pendingRows, workItem.SheetName);
                            pendingRows = 0;
                        }
                    });

                if (pendingRows > 0)
                {
                    tracker.AddRows(pendingRows, workItem.SheetName);
                }
            }

            tracker.MarkSheetCompleted(workItem.SheetName);
        }

        private static string ResolveOutputDirectory(string excelFilePath, string tsvDirectoryPath)
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

        private void ClearSelectedCells()
        {
            if (dataGridView1.SelectedCells == null || dataGridView1.SelectedCells.Count == 0)
            {
                return;
            }

            foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
            {
                if (cell == null || cell.ReadOnly)
                {
                    continue;
                }

                var row = cell.OwningRow;
                if (row != null && row.IsNewRow)
                {
                    continue;
                }

                cell.Value = string.Empty;
            }
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

        private static void WriteWorksheetAsTsv(
            WorksheetPart worksheetPart,
            IReadOnlyList<string> sharedStringValues,
            string outputFilePath,
            CancellationToken token,
            Action<long> onRowsWritten)
        {
            using (var writer = new StreamWriter(outputFilePath, false, new UTF8Encoding(false)))
            using (var reader = OpenXmlReader.Create(worksheetPart))
            {
                uint expectedRowIndex = 1;

                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();

                    if (!reader.IsStartElement || reader.ElementType != typeof(Row))
                    {
                        continue;
                    }

                    var row = (Row)reader.LoadCurrentElement();
                    var rowIndex = row.RowIndex != null ? row.RowIndex.Value : expectedRowIndex;

                    if (rowIndex > expectedRowIndex)
                    {
                        var blankRowCount = (long)(rowIndex - expectedRowIndex);
                        WriteBlankLines(writer, blankRowCount, token);
                        onRowsWritten(blankRowCount);
                    }

                    var rowValues = ReadRowValues(sharedStringValues, row);
                    writer.Write(string.Join("\t", rowValues.Select(EscapeTsvValue)));
                    writer.Write("\r\n");
                    onRowsWritten(1);

                    expectedRowIndex = rowIndex + 1;
                }
            }
        }

        private static void WriteBlankLines(StreamWriter writer, long count, CancellationToken token)
        {
            if (count <= 0)
            {
                return;
            }

            const int chunkSize = 4096;
            for (var i = 0L; i < count; i++)
            {
                if (i % chunkSize == 0)
                {
                    token.ThrowIfCancellationRequested();
                }

                writer.Write("\r\n");
            }
        }

        private static List<string> ReadRowValues(IReadOnlyList<string> sharedStringValues, Row row)
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

                values.Add(GetCellValue(sharedStringValues, cell));
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

        private static string GetCellValue(IReadOnlyList<string> sharedStringValues, Cell cell)
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
                return GetSharedStringValue(sharedStringValues, rawValue);
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

        private static string GetSharedStringValue(IReadOnlyList<string> sharedStringValues, string indexText)
        {
            int index;
            if (!int.TryParse(indexText, out index))
            {
                return string.Empty;
            }

            if (index < 0 || index >= sharedStringValues.Count)
            {
                return string.Empty;
            }

            return sharedStringValues[index] ?? string.Empty;
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

        private void Form1_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = string.Format("{0}  ver {1}.{2}.{3}", Text, version.Major, version.Minor, version.Build);
        }

        private sealed class ConversionRequest
        {
            public string ExcelFilePath { get; set; }

            public string OutputDirectory { get; set; }

            public List<SheetMapping> Mappings { get; set; }
        }

        private sealed class ConversionPlan
        {
            public List<string> SharedStringValues { get; set; }

            public List<SheetWorkItem> WorkItems { get; set; }

            public long TotalEstimatedRows { get; set; }
        }

        private sealed class SheetWorkItem
        {
            public string SheetName { get; set; }

            public string RelationshipId { get; set; }

            public string OutputFilePath { get; set; }

            public long EstimatedRows { get; set; }
        }

        private sealed class SheetMapping
        {
            public string SheetName { get; set; }

            public string TsvFileName { get; set; }
        }

        private sealed class ConversionProgress
        {
            public int TotalSheets { get; set; }

            public int CompletedSheets { get; set; }

            public long TotalEstimatedRows { get; set; }

            public long ProcessedRows { get; set; }

            public string Message { get; set; }
        }

        private sealed class ProgressTracker
        {
            private readonly long _totalEstimatedRows;
            private readonly int _totalSheets;
            private readonly IProgress<ConversionProgress> _progress;
            private long _processedRows;
            private int _completedSheets;

            public ProgressTracker(long totalEstimatedRows, int totalSheets, IProgress<ConversionProgress> progress)
            {
                _totalEstimatedRows = Math.Max(1, totalEstimatedRows);
                _totalSheets = Math.Max(1, totalSheets);
                _progress = progress;
            }

            public void AddRows(long rows, string message)
            {
                if (rows <= 0)
                {
                    return;
                }

                var processed = Interlocked.Add(ref _processedRows, rows);
                ReportInternal(processed, Volatile.Read(ref _completedSheets), message);
            }

            public void MarkSheetCompleted(string sheetName)
            {
                var completed = Interlocked.Increment(ref _completedSheets);
                var processed = Interlocked.Read(ref _processedRows);
                ReportInternal(processed, completed, sheetName + " 完了");
            }

            public void Report(string message)
            {
                var processed = Interlocked.Read(ref _processedRows);
                var completed = Volatile.Read(ref _completedSheets);
                ReportInternal(processed, completed, message);
            }

            private void ReportInternal(long processedRows, int completedSheets, string message)
            {
                if (_progress == null)
                {
                    return;
                }

                _progress.Report(new ConversionProgress
                {
                    TotalSheets = _totalSheets,
                    CompletedSheets = Math.Max(0, Math.Min(_totalSheets, completedSheets)),
                    TotalEstimatedRows = _totalEstimatedRows,
                    ProcessedRows = Math.Max(0, Math.Min(_totalEstimatedRows, processedRows)),
                    Message = message ?? string.Empty
                });
            }
        }
    }
}
