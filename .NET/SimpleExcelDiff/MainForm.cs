using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Color = System.Drawing.Color;

namespace SimpleExcelDiff
{
    public partial class MainForm : Form
    {
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();
            settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleExcelDiff_Settings.json");
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
            lblStatus.Text = "準備完了";
            this.txtPathSrc.AllowDrop = true;
            this.txtPathSrc.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.txtPathSrc.DragDrop += new DragEventHandler(textBox_DragDrop);
            this.txtPathDst.AllowDrop = true;
            this.txtPathDst.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.txtPathDst.DragDrop += new DragEventHandler(textBox_DragDrop);
            
            this.btnColorSelector.Click += new System.EventHandler(this.btnColorSelector_Click);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void LoadSettings()
        {
            if (!File.Exists(settingsFilePath)) return;
            try
            {
                using (var fs = new FileStream(settingsFilePath, FileMode.Open, FileAccess.Read))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SimpleExcelDiffSettings));
                    var settings = (SimpleExcelDiffSettings)serializer.ReadObject(fs);
                    txtPathSrc.Text = settings.PathSrc;
                    txtPathDst.Text = settings.PathDst;
                    chkEnableSubDir.Checked = settings.EnableSubDir;
                    picCellColor.BackColor = Color.FromArgb(settings.HighlightColorArgb);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettings()
        {
            try
            {
                var settings = new SimpleExcelDiffSettings
                {
                    PathSrc = txtPathSrc.Text,
                    PathDst = txtPathDst.Text,
                    EnableSubDir = chkEnableSubDir.Checked,
                    HighlightColorArgb = picCellColor.BackColor.ToArgb()
                };
                using (var fs = new FileStream(settingsFilePath, FileMode.Create, FileAccess.Write))
                {
                    var serializer = new DataContractJsonSerializer(typeof(SimpleExcelDiffSettings));
                    serializer.WriteObject(fs, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            else e.Effect = DragDropEffects.None;
        }

        private void textBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length > 0)
            {
                ((TextBox)sender).Text = files[0];
            }
        }


        private void btnBrowseSrc_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPathSrc.Text = fbd.SelectedPath;
                }
            }
        }

        private void btnBrowseDst_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtPathDst.Text = fbd.SelectedPath;
                }
            }
        }

        private async void btnProcess_Click(object sender, EventArgs e)
        {
            if (chkDrawCell.Checked)
            {
                var result = MessageBox.Show(
                    "差分があるセルの背景色やシートのタブ色を変更します。\r\nファイルが更新されますが処理を続けますか？",
                    "確認",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Cancel)
                {
                    lblStatus.Text = "処理がキャンセルされました。";
                    return;
                }
            }

            lblStatus.Text = "処理中...";
            this.Enabled = false;
            dataGridView1.DataSource = null;

            var pathSrc = txtPathSrc.Text;
            var pathDst = txtPathDst.Text;
            var includeSubDir = chkEnableSubDir.Checked;
            var highlightColor = picCellColor.BackColor;

            try
            {
                var diffResults = await Task.Run(() => FindDifferences(pathSrc, pathDst, includeSubDir, chkDrawCell.Checked, highlightColor));
                DisplayResults(diffResults);
                lblStatus.Text = $"処理完了: {diffResults.Count}件の差分/エラーを検出しました。";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "エラーが発生しました。";
                MessageBox.Show($"処理中にエラーが発生しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Enabled = true;
            }
        }

        private List<DiffResult> FindDifferences(string pathSrc, string pathDst, bool includeSubDir, bool applyColoring, Color highlightColor)
        {
            var allDiffResults = new List<DiffResult>();
            
            if (File.Exists(pathSrc) && File.Exists(pathDst))
            {
                 allDiffResults.AddRange(CompareExcelFiles(pathSrc, pathDst));
            }
            else if (Directory.Exists(pathSrc) && Directory.Exists(pathDst))
            {
                var searchOption = includeSubDir ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var filesSrc = GetExcelFiles(pathSrc, searchOption);
                var filesDst = GetExcelFiles(pathDst, searchOption);

                var fileMapSrc = filesSrc.ToDictionary(f => f.Substring(pathSrc.Length), f => f, StringComparer.OrdinalIgnoreCase);
                var fileMapDst = filesDst.ToDictionary(f => f.Substring(pathDst.Length), f => f, StringComparer.OrdinalIgnoreCase);

                foreach (var relativePath in fileMapSrc.Keys.Intersect(fileMapDst.Keys, StringComparer.OrdinalIgnoreCase))
                {
                    allDiffResults.AddRange(CompareExcelFiles(fileMapSrc[relativePath], fileMapDst[relativePath]));
                }
            }
            else
            {
                 throw new FileNotFoundException("指定された比較元または比較先のパスが見つかりません。");
            }
            
            int initialId = 1;
            allDiffResults.ForEach(r => r.Id = initialId++);

            if (applyColoring && allDiffResults.Any(d => d.DiffType != DiffType.WriteError))
            {
                var writeErrors = ApplyColoringToDifferences(allDiffResults, highlightColor);
                allDiffResults.AddRange(writeErrors);
            }
            
            var displayResults = allDiffResults
                .Where(r => r.DiffType == DiffType.WriteError) 
                .Union(allDiffResults
                    .Where(r => r.DiffType != DiffType.WriteError)
                    .GroupBy(r => new { r.FileNameSrc, r.FileNameDst, r.SheetName }) 
                    .Select(g => g.OrderBy(d => d.Id).First()))
                .OrderBy(r => r.Id)
                .ToList();

            int finalId = 1;
            displayResults.ForEach(r => r.Id = finalId++);

            return displayResults;
        }

        private IEnumerable<string> GetExcelFiles(string path, SearchOption searchOption)
        {
            if (File.Exists(path) && (path.EndsWith(".xlsx") || path.EndsWith(".xlsm"))) return new[] { path };
            if (Directory.Exists(path)) return Directory.EnumerateFiles(path, "*.xlsx", searchOption).Union(Directory.EnumerateFiles(path, "*.xlsm", searchOption));
            return Enumerable.Empty<string>();
        }
        
        private List<DiffResult> CompareExcelFiles(string fileSrc, string fileDst)
        {
            var results = new List<DiffResult>();
            try
            {
                using (var docSrc = SpreadsheetDocument.Open(fileSrc, false))
                using (var docDst = SpreadsheetDocument.Open(fileDst, false))
                {
                    var wbPartSrc = docSrc.WorkbookPart;
                    var wbPartDst = docDst.WorkbookPart;

                    if (wbPartSrc?.Workbook == null || wbPartDst?.Workbook == null)
                    {
                        results.Add(new DiffResult { DiffType = DiffType.WriteError, ErrorMessage = "ファイル形式エラー: Excelファイルの内部構造が不正です（Workbookが見つかりません）。" });
                        results.ForEach(r => {
                            r.FolderPathSrc = Path.GetDirectoryName(fileSrc);
                            r.FileNameSrc = Path.GetFileName(fileSrc);
                            r.FolderPathDst = Path.GetDirectoryName(fileDst);
                            r.FileNameDst = Path.GetFileName(fileDst);
                        });
                        return results;
                    }

                    var sheetsSrc = wbPartSrc.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToDictionary(s => s.Name.Value, s => s.Id.Value);
                    var sheetsDst = wbPartDst.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToDictionary(s => s.Name.Value, s => s.Id.Value);

                    foreach (var sheetName in sheetsSrc.Keys.Except(sheetsDst.Keys)) results.Add(new DiffResult { DiffType = DiffType.SheetMissingInDst, SheetName = sheetName });
                    foreach (var sheetName in sheetsDst.Keys.Except(sheetsSrc.Keys)) results.Add(new DiffResult { DiffType = DiffType.SheetMissingInSrc, SheetName = sheetName });
                    
                    foreach (var sheetName in sheetsSrc.Keys.Intersect(sheetsDst.Keys))
                    {
                        var wsPartSrc = (WorksheetPart)wbPartSrc.GetPartById(sheetsSrc[sheetName]);
                        var wsPartDst = (WorksheetPart)wbPartDst.GetPartById(sheetsDst[sheetName]);
                        results.AddRange(CompareSheets(wsPartSrc, wsPartDst, wbPartSrc, wbPartDst, sheetName));
                    }
                }
            }
            catch (IOException ex)
            {
                results.Add(new DiffResult { DiffType = DiffType.WriteError, ErrorMessage = $"ファイルアクセスエラー: ファイルが他のプロセスで使用中の可能性があります。({ex.Message})" });
            }
            catch (OpenXmlPackageException ex)
            {
                results.Add(new DiffResult { DiffType = DiffType.WriteError, ErrorMessage = $"ファイル形式エラー: ファイルが破損しているか、サポートされていない形式（パスワード付き、古いExcel形式等）の可能性があります。({ex.Message})" });
            }
            catch (Exception ex)
            {
                results.Add(new DiffResult { DiffType = DiffType.WriteError, ErrorMessage = $"予期せぬエラー: {ex.Message}" });
            }

            results.ForEach(r => {
                r.FolderPathSrc = Path.GetDirectoryName(fileSrc);
                r.FileNameSrc = Path.GetFileName(fileSrc);
                r.FolderPathDst = Path.GetDirectoryName(fileDst);
                r.FileNameDst = Path.GetFileName(fileDst);
            });
            return results;
        }

        private List<DiffResult> CompareSheets(WorksheetPart wsPartSrc, WorksheetPart wsPartDst, WorkbookPart wbPartSrc, WorkbookPart wbPartDst, string sheetName)
        {
            var sheetResults = new List<DiffResult>();
            var rowsSrc = BuildSheetRows(wsPartSrc, wbPartSrc);
            var rowsDst = BuildSheetRows(wsPartDst, wbPartDst);

            var rowOps = CalculateDiffOperations(rowsSrc, rowsDst, RowsAreEqual);

            int opIndex = 0;
            while (opIndex < rowOps.Count)
            {
                var op = rowOps[opIndex];
                switch (op.Type)
                {
                    case SequenceDiffOperationType.Match:
                        CompareRowCells(rowsSrc[op.IndexSrc], rowsDst[op.IndexDst], sheetName, sheetResults);
                        opIndex++;
                        break;
                    case SequenceDiffOperationType.Delete:
                        if (opIndex + 1 < rowOps.Count && rowOps[opIndex + 1].Type == SequenceDiffOperationType.Insert)
                        {
                            var pairedInsert = rowOps[opIndex + 1];
                            CompareRowCells(rowsSrc[op.IndexSrc], rowsDst[pairedInsert.IndexDst], sheetName, sheetResults);
                            opIndex += 2;
                        }
                        else
                        {
                            AddRowCellsAsDiff(rowsSrc[op.IndexSrc], null, sheetName, sheetResults);
                            opIndex++;
                        }
                        break;
                    case SequenceDiffOperationType.Insert:
                        if (opIndex + 1 < rowOps.Count && rowOps[opIndex + 1].Type == SequenceDiffOperationType.Delete)
                        {
                            var pairedDelete = rowOps[opIndex + 1];
                            CompareRowCells(rowsSrc[pairedDelete.IndexSrc], rowsDst[op.IndexDst], sheetName, sheetResults);
                            opIndex += 2;
                        }
                        else
                        {
                            AddRowCellsAsDiff(null, rowsDst[op.IndexDst], sheetName, sheetResults);
                            opIndex++;
                        }
                        break;
                }
            }

            return sheetResults;
        }

        private List<SheetRow> BuildSheetRows(WorksheetPart worksheetPart, WorkbookPart workbookPart)
        {
            var rows = new List<SheetRow>();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return rows;

            foreach (var row in sheetData.Elements<Row>())
            {
                var sheetRow = new SheetRow(row.RowIndex?.Value ?? 0);
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellRef = cell.CellReference?.Value;
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    var columnIndex = GetColumnIndex(cellRef);
                    sheetRow.Cells.Add(new SheetCell
                    {
                        ColumnIndex = columnIndex,
                        CellReference = cellRef,
                        Value = GetCellValue(cell, workbookPart)
                    });
                }

                if (sheetRow.Cells.Count == 0) continue;
                sheetRow.Cells.Sort((a, b) => a.ColumnIndex.CompareTo(b.ColumnIndex));
                rows.Add(sheetRow);
            }

            return rows;
        }

        private bool RowsAreEqual(SheetRow left, SheetRow right)
        {
            if (left.Cells.Count != right.Cells.Count) return false;
            for (int i = 0; i < left.Cells.Count; i++)
            {
                if (!string.Equals(left.Cells[i].Value, right.Cells[i].Value, StringComparison.Ordinal)) return false;
            }
            return true;
        }

        private void CompareRowCells(SheetRow rowSrc, SheetRow rowDst, string sheetName, List<DiffResult> results)
        {
            var cellsSrc = rowSrc?.Cells ?? new List<SheetCell>();
            var cellsDst = rowDst?.Cells ?? new List<SheetCell>();

            var cellOps = CalculateDiffOperations(cellsSrc, cellsDst, (a, b) => string.Equals(a.Value, b.Value, StringComparison.Ordinal));

            int opIndex = 0;
            while (opIndex < cellOps.Count)
            {
                var op = cellOps[opIndex];
                switch (op.Type)
                {
                    case SequenceDiffOperationType.Match:
                        opIndex++;
                        break;
                    case SequenceDiffOperationType.Delete:
                        if (opIndex + 1 < cellOps.Count && cellOps[opIndex + 1].Type == SequenceDiffOperationType.Insert)
                        {
                            var pairedInsert = cellOps[opIndex + 1];
                            AddCellDiff(rowSrc.Cells[op.IndexSrc], rowDst.Cells[pairedInsert.IndexDst], sheetName, results);
                            opIndex += 2;
                        }
                        else
                        {
                            AddCellDiff(rowSrc.Cells[op.IndexSrc], null, sheetName, results);
                            opIndex++;
                        }
                        break;
                    case SequenceDiffOperationType.Insert:
                        if (opIndex + 1 < cellOps.Count && cellOps[opIndex + 1].Type == SequenceDiffOperationType.Delete)
                        {
                            var pairedDelete = cellOps[opIndex + 1];
                            AddCellDiff(rowSrc.Cells[pairedDelete.IndexSrc], rowDst.Cells[op.IndexDst], sheetName, results);
                            opIndex += 2;
                        }
                        else
                        {
                            AddCellDiff(null, rowDst.Cells[op.IndexDst], sheetName, results);
                            opIndex++;
                        }
                        break;
                }
            }
        }

        private void AddRowCellsAsDiff(SheetRow rowSrc, SheetRow rowDst, string sheetName, List<DiffResult> results)
        {
            if (rowSrc != null)
            {
                foreach (var cell in rowSrc.Cells)
                {
                    AddCellDiff(cell, null, sheetName, results);
                }
            }

            if (rowDst != null)
            {
                foreach (var cell in rowDst.Cells)
                {
                    AddCellDiff(null, cell, sheetName, results);
                }
            }
        }

        private void AddCellDiff(SheetCell cellSrc, SheetCell cellDst, string sheetName, List<DiffResult> results)
        {
            if (cellSrc == null && cellDst == null) return;

            var cellAddress = cellDst?.CellReference ?? cellSrc?.CellReference ?? "";

            results.Add(new DiffResult
            {
                DiffType = DiffType.CellValueMismatch,
                SheetName = sheetName,
                CellAddress = cellAddress,
                CellValueSrc = cellSrc?.Value ?? "",
                CellValueDst = cellDst?.Value ?? ""
            });
        }

        private List<SequenceDiffOperation> CalculateDiffOperations<T>(IList<T> source, IList<T> destination, Func<T, T, bool> comparer)
        {
            int m = source.Count;
            int n = destination.Count;
            var dp = new int[m + 1, n + 1];

            for (int i = m - 1; i >= 0; i--)
            {
                for (int j = n - 1; j >= 0; j--)
                {
                    if (comparer(source[i], destination[j]))
                    {
                        dp[i, j] = dp[i + 1, j + 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = Math.Max(dp[i + 1, j], dp[i, j + 1]);
                    }
                }
            }

            var operations = new List<SequenceDiffOperation>();
            int srcIndex = 0;
            int dstIndex = 0;

            while (srcIndex < m && dstIndex < n)
            {
                if (comparer(source[srcIndex], destination[dstIndex]))
                {
                    operations.Add(new SequenceDiffOperation(SequenceDiffOperationType.Match, srcIndex, dstIndex));
                    srcIndex++;
                    dstIndex++;
                }
                else if (dp[srcIndex + 1, dstIndex] >= dp[srcIndex, dstIndex + 1])
                {
                    operations.Add(new SequenceDiffOperation(SequenceDiffOperationType.Delete, srcIndex, dstIndex));
                    srcIndex++;
                }
                else
                {
                    operations.Add(new SequenceDiffOperation(SequenceDiffOperationType.Insert, srcIndex, dstIndex));
                    dstIndex++;
                }
            }

            while (srcIndex < m)
            {
                operations.Add(new SequenceDiffOperation(SequenceDiffOperationType.Delete, srcIndex, dstIndex));
                srcIndex++;
            }

            while (dstIndex < n)
            {
                operations.Add(new SequenceDiffOperation(SequenceDiffOperationType.Insert, srcIndex, dstIndex));
                dstIndex++;
            }

            return operations;
        }

        private int GetColumnIndex(string cellReference)
        {
            int index = 0;
            foreach (var ch in cellReference.ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') break;
                index = index * 26 + (ch - 'A' + 1);
            }
            return index;
        }

        private sealed class SheetRow
        {
            public SheetRow(uint rowIndex)
            {
                RowIndex = rowIndex;
            }

            public uint RowIndex { get; }
            public List<SheetCell> Cells { get; } = new List<SheetCell>();
        }

        private sealed class SheetCell
        {
            public int ColumnIndex { get; set; }
            public string CellReference { get; set; }
            public string Value { get; set; }
        }

        private enum SequenceDiffOperationType
        {
            Match,
            Delete,
            Insert
        }

        private readonly struct SequenceDiffOperation
        {
            public SequenceDiffOperation(SequenceDiffOperationType type, int indexSrc, int indexDst)
            {
                Type = type;
                IndexSrc = indexSrc;
                IndexDst = indexDst;
            }

            public SequenceDiffOperationType Type { get; }
            public int IndexSrc { get; }
            public int IndexDst { get; }
        }

        private string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            if (cell?.CellValue == null) return "";
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) return wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable.ElementAt(int.Parse(value)).InnerText ?? value;
            return value;
        }

        private void DisplayResults(List<DiffResult> results)
        {
            var dt = new DataTable();
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("差分種別", typeof(string));
            dt.Columns.Add("比較元フォルダ", typeof(string));
            dt.Columns.Add("比較元ファイル", typeof(string));
            dt.Columns.Add("比較元シート", typeof(string));
            dt.Columns.Add("比較元セル", typeof(string));
            dt.Columns.Add("比較元値/エラー内容", typeof(string));
            dt.Columns.Add("比較先フォルダ", typeof(string));
            dt.Columns.Add("比較先ファイル", typeof(string));
            dt.Columns.Add("比較先シート", typeof(string));
            dt.Columns.Add("比較先セル", typeof(string));
            dt.Columns.Add("比較先値", typeof(string));
            dt.Columns.Add("DiffTypeEnum", typeof(DiffType));

            foreach (var r in results)
            {
                string diffTypeText = "";
                switch (r.DiffType)
                {
                    case DiffType.CellValueMismatch:
                        diffTypeText = "値の不一致";
                        dt.Rows.Add(r.Id, diffTypeText, r.FolderPathSrc, r.FileNameSrc, r.SheetName, r.CellAddress, r.CellValueSrc, r.FolderPathDst, r.FileNameDst, r.SheetName, r.CellAddress, r.CellValueDst, r.DiffType);
                        break;
                    case DiffType.SheetMissingInDst:
                        diffTypeText = "シート不足(比較先)";
                        dt.Rows.Add(r.Id, diffTypeText, r.FolderPathSrc, r.FileNameSrc, r.SheetName, null, null, r.FolderPathDst, r.FileNameDst, null, null, null, r.DiffType);
                        break;
                    case DiffType.SheetMissingInSrc:
                         diffTypeText = "シート不足(比較元)";
                        dt.Rows.Add(r.Id, diffTypeText, r.FolderPathSrc, r.FileNameSrc, null, null, null, r.FolderPathDst, r.FileNameDst, r.SheetName, null, null, r.DiffType);
                        break;
                    case DiffType.WriteError:
                        diffTypeText = "書き込みエラー";
                        var errorPath = !string.IsNullOrEmpty(r.FileNameSrc) ? Path.Combine(r.FolderPathSrc, r.FileNameSrc) : Path.Combine(r.FolderPathDst, r.FileNameDst);
                        dt.Rows.Add(r.Id, diffTypeText, r.FolderPathSrc, r.FileNameSrc, null, null, $"[{Path.GetFileName(errorPath)}] {r.ErrorMessage}", r.FolderPathDst, r.FileNameDst, null, null, null, r.DiffType);
                        break;
                }
            }
            
            dataGridView1.DataSource = dt;
            dataGridView1.Columns["DiffTypeEnum"].Visible = false;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;
                var diffType = (DiffType)row.Cells["DiffTypeEnum"].Value;
                switch (diffType)
                {
                    case DiffType.CellValueMismatch:
                        row.Cells["比較元値/エラー内容"].Style.BackColor = System.Drawing.Color.Yellow;
                        row.Cells["比較先値"].Style.BackColor = System.Drawing.Color.Yellow;
                        break;
                    case DiffType.SheetMissingInDst:
                        row.Cells["比較元シート"].Style.BackColor = System.Drawing.Color.Yellow;
                        break;
                    case DiffType.SheetMissingInSrc:
                        row.Cells["比較先シート"].Style.BackColor = System.Drawing.Color.Yellow;
                        break;
                    case DiffType.WriteError:
                        row.Cells["差分種別"].Style.BackColor = System.Drawing.Color.Red;
                        row.Cells["差分種別"].Style.ForeColor = System.Drawing.Color.White;
                        row.Cells["比較元値/エラー内容"].Style.BackColor = System.Drawing.Color.MistyRose;
                        break;
                }
            }
            dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void btnColorSelector_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = picCellColor.BackColor;
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    picCellColor.BackColor = colorDialog.Color;
                }
            }
        }
        
        private List<DiffResult> ApplyColoringToDifferences(List<DiffResult> diffs, Color highlightColor)
        {
            var writeErrors = new List<DiffResult>();
            var groupedByFile = diffs.GroupBy(d => new { d.FolderPathSrc, d.FileNameSrc, d.FolderPathDst, d.FileNameDst });

            foreach (var fileGroup in groupedByFile)
            {
                var srcPath = Path.Combine(fileGroup.Key.FolderPathSrc, fileGroup.Key.FileNameSrc);
                try
                {
                    using (var doc = SpreadsheetDocument.Open(srcPath, true))
                    {
                        uint styleIndex = CreateAndGetStyleIndex(doc, highlightColor);
                        foreach (var diff in fileGroup)
                        {
                            if (diff.DiffType == DiffType.CellValueMismatch)
                            {
                                SetCellFill(doc, diff.SheetName, diff.CellAddress, styleIndex);
                            }
                            // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
                            // ★★★ 安定動作のため、エラーの根本原因であるシートタブの色変更を無効化 ★★★
                            // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
                            // else if (diff.DiffType == DiffType.SheetMissingInDst)
                            // {
                            //     SetSheetTabColor(doc, diff.SheetName, highlightColor);
                            // }
                        }
                        doc.Save();
                    }
                }
                catch (Exception ex)
                {
                    writeErrors.Add(new DiffResult { DiffType = DiffType.WriteError, FolderPathSrc = fileGroup.Key.FolderPathSrc, FileNameSrc = fileGroup.Key.FileNameSrc, ErrorMessage = $"比較元ファイル書き込み不可: {ex.Message}" });
                }

                var dstPath = Path.Combine(fileGroup.Key.FolderPathDst, fileGroup.Key.FileNameDst);
                try
                {
                    using (var doc = SpreadsheetDocument.Open(dstPath, true))
                    {
                        uint styleIndex = CreateAndGetStyleIndex(doc, highlightColor);
                        foreach (var diff in fileGroup)
                        {
                           if (diff.DiffType == DiffType.CellValueMismatch)
                           {
                                SetCellFill(doc, diff.SheetName, diff.CellAddress, styleIndex);
                           }
                           // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
                           // ★★★ 安定動作のため、エラーの根本原因であるシートタブの色変更を無効化 ★★★
                           // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
                           // else if (diff.DiffType == DiffType.SheetMissingInSrc)
                           // {
                           //      SetSheetTabColor(doc, diff.SheetName, highlightColor);
                           // }
                        }
                        doc.Save();
                    }
                }
                catch (Exception ex)
                {
                    writeErrors.Add(new DiffResult { DiffType = DiffType.WriteError, FolderPathDst = fileGroup.Key.FolderPathDst, FileNameDst = fileGroup.Key.FileNameDst, ErrorMessage = $"比較先ファイル書き込み不可: {ex.Message}" });
                }
            }
            return writeErrors;
        }
        
        private void SetSheetTabColor(SpreadsheetDocument doc, string sheetName, Color color)
        {
            var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null) return;
            
            var sheetProperties = sheet.Elements<SheetProperties>().FirstOrDefault();
            if (sheetProperties == null)
            {
                sheetProperties = new SheetProperties();
                sheet.Append(sheetProperties);
            }
            
            var tabColorElement = sheetProperties.Elements<TabColor>().FirstOrDefault();
            if (tabColorElement != null)
            {
                tabColorElement.Remove();
            }
            
            sheetProperties.Append(new TabColor() { Rgb = new HexBinaryValue() { Value = $"{color.R:X2}{color.G:X2}{color.B:X2}" }});
        }
        
        private void SetCellFill(SpreadsheetDocument doc, string sheetName, string cellAddress, uint styleIndex)
        {
            var sheet = doc.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null) return;

            var wsPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var cell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellAddress);
            if (cell == null) return;

            cell.StyleIndex = styleIndex;
        }
        
        private uint CreateAndGetStyleIndex(SpreadsheetDocument doc, Color color)
        {
            var stylesPart = doc.WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null)
            {
                stylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

            if (stylesPart.Stylesheet == null)
            {
                stylesPart.Stylesheet = new Stylesheet();
            }
            var stylesheet = stylesPart.Stylesheet;

            if (stylesheet.Elements<Fonts>().Count() == 0)
            {
                stylesheet.Append(new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()));
            }
            if (stylesheet.Elements<Fills>().Count() == 0)
            {
                stylesheet.Append(new Fills(
                    new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } },
                    new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }
                ));
            }
            if (stylesheet.Elements<Borders>().Count() == 0)
            {
                stylesheet.Append(new Borders(new Border()));
            }
            if (stylesheet.Elements<CellStyleFormats>().Count() == 0)
            {
                stylesheet.Append(new CellStyleFormats(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }));
            }
            if (stylesheet.Elements<CellFormats>().Count() == 0)
            {
                 stylesheet.Append(new CellFormats(new CellFormat()));
            }

            var newFill = new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = new HexBinaryValue() { Value = $"{color.R:X2}{color.G:X2}{color.B:X2}" } }
                ) { PatternType = PatternValues.Solid }
            );

            Fills fills = stylesheet.Elements<Fills>().First();
            uint fillIndex = 0;
            var existingFill = fills.Elements<Fill>().FirstOrDefault(f => f.OuterXml == newFill.OuterXml);
            if (existingFill != null)
            {
                fillIndex = (uint)Array.IndexOf(fills.Elements<Fill>().ToArray(), existingFill);
            }
            else
            {
                fills.Append(newFill);
                fillIndex = (uint)fills.Count() - 1;
            }

            var newCellFormat = new CellFormat {
                NumberFormatId = 0, FontId = 0, BorderId = 0, FormatId = 0,
                FillId = fillIndex, 
                ApplyFill = true 
            };
            
            CellFormats cellFormats = stylesheet.Elements<CellFormats>().First();
            uint styleIndex = 0;
            var existingFormat = cellFormats.Elements<CellFormat>().FirstOrDefault(cf => cf.OuterXml == newCellFormat.OuterXml);
            if (existingFormat != null)
            {
                styleIndex = (uint)Array.IndexOf(cellFormats.Elements<CellFormat>().ToArray(), existingFormat);
            }
            else
            {
                cellFormats.Append(newCellFormat);
                styleIndex = (uint)cellFormats.Count() - 1;
            }
            
            return styleIndex;
        }

        private void checkBox1_CheckedChanged ( object sender, EventArgs e )
        {
        }
    }
    
    [DataContract]
    internal class SimpleExcelDiffSettings
    {
        [DataMember] public string PathSrc { get; set; }
        [DataMember] public string PathDst { get; set; }
        [DataMember] public bool EnableSubDir { get; set; }
        [DataMember] public int HighlightColorArgb { get; set; } = Color.Yellow.ToArgb();
    }
    
    internal enum DiffType { CellValueMismatch, SheetMissingInDst, SheetMissingInSrc, WriteError }

    internal class DiffResult
    {
        public int Id { get; set; }
        public DiffType DiffType { get; set; }
        public string SheetName { get; set; }
        public string CellAddress { get; set; }
        public string CellValueSrc { get; set; }
        public string CellValueDst { get; set; }
        public string FolderPathSrc { get; set; }
        public string FileNameSrc { get; set; }
        public string FolderPathDst { get; set; }
        public string FileNameDst { get; set; }
        public string ErrorMessage { get; set; }
    }
}
