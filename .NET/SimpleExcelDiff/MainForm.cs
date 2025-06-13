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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SimpleExcelDiff
{
    public partial class MainForm : Form
    {
        private readonly string settingsFilePath;

        public MainForm()
        {
            InitializeComponent();

            // 設定ファイルのパスを定義
            settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleExcelDiff_Settings.json");

            // イベントハンドラを登録
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;

            lblStatus.Text = "準備完了";

            // 両方のテキストボックスでドラッグ＆ドロップを有効化
            this.txtPathSrc.AllowDrop = true;
            this.txtPathSrc.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.txtPathSrc.DragDrop += new DragEventHandler(textBox_DragDrop);

            this.txtPathDst.AllowDrop = true;
            this.txtPathDst.DragEnter += new DragEventHandler(textBox_DragEnter);
            this.txtPathDst.DragDrop += new DragEventHandler(textBox_DragDrop);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // バージョン情報をタイトルに表示
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
                    EnableSubDir = chkEnableSubDir.Checked
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
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
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
            lblStatus.Text = "処理中...";
            this.Enabled = false;
            // ★★★ UIに合わせて dgvSrc と dgvDst をクリア
            dgvSrc.DataSource = null;
            dgvDst.DataSource = null;

            var pathSrc = txtPathSrc.Text;
            var pathDst = txtPathDst.Text;
            var includeSubDir = chkEnableSubDir.Checked;

            try
            {
                var diffResults = await Task.Run(() => FindDifferences(pathSrc, pathDst, includeSubDir));
                DisplayResults(diffResults); // 新しい表示メソッドを呼び出し
                lblStatus.Text = $"処理完了: {diffResults.Count}件の差分を検出しました。";
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

        /// <summary>
        /// 差分を検索するメインロジック
        /// </summary>
        private List<DiffResult> FindDifferences(string pathSrc, string pathDst, bool includeSubDir)
        {
            var diffResults = new List<DiffResult>();
            int id = 1;

            // 仕様変更②: ファイルが直接指定された場合の処理
            if (File.Exists(pathSrc) && File.Exists(pathDst))
            {
                var results = CompareExcelFiles(pathSrc, pathDst);
                foreach (var res in results)
                {
                    res.Id = id++;
                    diffResults.Add(res);
                }
            }
            // 仕様変更①: 従来通りフォルダが指定された場合の処理
            else if (Directory.Exists(pathSrc) && Directory.Exists(pathDst))
            {
                var searchOption = includeSubDir ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var filesSrc = GetExcelFiles(pathSrc, searchOption);
                var filesDst = GetExcelFiles(pathDst, searchOption);

                var fileMapSrc = filesSrc.ToDictionary(f => Path.GetFileName(f), f => f);
                var fileMapDst = filesDst.ToDictionary(f => Path.GetFileName(f), f => f);

                // ファイル名が共通のものを比較
                foreach (var fileName in fileMapSrc.Keys.Intersect(fileMapDst.Keys))
                {
                    var results = CompareExcelFiles(fileMapSrc[fileName], fileMapDst[fileName]);
                    foreach (var res in results)
                    {
                        res.Id = id++;
                        diffResults.Add(res);
                    }
                }
            }
            else
            {
                 // どちらかのパス、あるいは両方が無効な場合
                 throw new FileNotFoundException("指定された比較元または比較先のパスが見つかりません。");
            }
            
            return diffResults;
        }

        private IEnumerable<string> GetExcelFiles(string path, SearchOption searchOption)
        {
            if (File.Exists(path) && (path.EndsWith(".xlsx") || path.EndsWith(".xlsm")))
            {
                return new[] { path };
            }
            if (Directory.Exists(path))
            {
                return Directory.EnumerateFiles(path, "*.xlsx", searchOption)
                                .Union(Directory.EnumerateFiles(path, "*.xlsm", searchOption));
            }
            return Enumerable.Empty<string>();
        }

        private List<DiffResult> CompareExcelFiles(string fileSrc, string fileDst)
        {
            var results = new List<DiffResult>();
            using (var docSrc = SpreadsheetDocument.Open(fileSrc, false))
            using (var docDst = SpreadsheetDocument.Open(fileDst, false))
            {
                var wbPartSrc = docSrc.WorkbookPart;
                var wbPartDst = docDst.WorkbookPart;

                var sheetsSrc = wbPartSrc.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToDictionary(s => s.Name.Value, s => s.Id.Value);
                var sheetsDst = wbPartDst.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToDictionary(s => s.Name.Value, s => s.Id.Value);

                foreach (var sheetName in sheetsSrc.Keys.Except(sheetsDst.Keys))
                {
                    results.Add(new DiffResult { DiffType = DiffType.SheetMissingInDst, SheetName = sheetName });
                }
                foreach (var sheetName in sheetsDst.Keys.Except(sheetsSrc.Keys))
                {
                    results.Add(new DiffResult { DiffType = DiffType.SheetMissingInSrc, SheetName = sheetName });
                }
                foreach (var sheetName in sheetsSrc.Keys.Intersect(sheetsDst.Keys))
                {
                    var wsPartSrc = (WorksheetPart)wbPartSrc.GetPartById(sheetsSrc[sheetName]);
                    var wsPartDst = (WorksheetPart)wbPartDst.GetPartById(sheetsDst[sheetName]);
                    var diff = CompareSheets(wsPartSrc, wsPartDst, wbPartSrc, wbPartDst);
                    if (diff != null)
                    {
                        diff.SheetName = sheetName;
                        results.Add(diff);
                    }
                }
            }
            // ファイル情報をここでまとめて付与
            results.ForEach(r => {
                r.FolderPathSrc = Path.GetDirectoryName(fileSrc);
                r.FileNameSrc = Path.GetFileName(fileSrc);
                r.FolderPathDst = Path.GetDirectoryName(fileDst);
                r.FileNameDst = Path.GetFileName(fileDst);
            });
            return results;
        }

        private DiffResult CompareSheets(WorksheetPart wsPartSrc, WorksheetPart wsPartDst, WorkbookPart wbPartSrc, WorkbookPart wbPartDst)
        {
            var cellsSrc = wsPartSrc.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference.Value, c => c);
            var cellsDst = wsPartDst.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference.Value, c => c);
            var allCellReferences = cellsSrc.Keys.Union(cellsDst.Keys).Distinct()
                .OrderBy(r => uint.Parse(System.Text.RegularExpressions.Regex.Match(r, @"\d+").Value))
                .ThenBy(r => System.Text.RegularExpressions.Regex.Match(r, @"[A-Z]+").Value);

            foreach (var cellRef in allCellReferences)
            {
                var valSrc = cellsSrc.ContainsKey(cellRef) ? GetCellValue(cellsSrc[cellRef], wbPartSrc) : "";
                var valDst = cellsDst.ContainsKey(cellRef) ? GetCellValue(cellsDst[cellRef], wbPartDst) : "";

                if (valSrc != valDst)
                {
                    return new DiffResult { DiffType = DiffType.CellValueMismatch, CellAddress = cellRef, CellValueSrc = valSrc, CellValueDst = valDst };
                }
            }
            return null;
        }

        private string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            if (cell?.CellValue == null) return "";
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()?.SharedStringTable.ElementAt(int.Parse(value)).InnerText ?? value;
            }
            return value;
        }

        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        // ★★★ DisplayResultsメソッドを2つの表に対応するように全面的に書き換え ★★★
        // ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★
        private void DisplayResults(List<DiffResult> results)
        {
            var dtSrc = new DataTable();
            var dtDst = new DataTable();
            var columns = new[] { "ID", "フォルダパス", "ファイル名", "シート名", "セル位置", "セル値" };
            foreach (var col in columns)
            {
                dtSrc.Columns.Add(col);
                dtDst.Columns.Add(col);
            }

            var resultsDict = results.ToDictionary(r => r.Id, r => r);

            foreach (var r in results)
            {
                DataRow drSrc = dtSrc.NewRow();
                DataRow drDst = dtDst.NewRow();

                switch (r.DiffType)
                {
                    case DiffType.CellValueMismatch:
                        drSrc.ItemArray = new object[] { r.Id, r.FolderPathSrc, r.FileNameSrc, r.SheetName, r.CellAddress, r.CellValueSrc };
                        drDst.ItemArray = new object[] { r.Id, r.FolderPathDst, r.FileNameDst, r.SheetName, r.CellAddress, r.CellValueDst };
                        dtSrc.Rows.Add(drSrc);
                        dtDst.Rows.Add(drDst);
                        break;
                    case DiffType.SheetMissingInDst: // 比較元にのみシートが存在
                        drSrc.ItemArray = new object[] { r.Id, r.FolderPathSrc, r.FileNameSrc, r.SheetName, r.CellAddress, r.CellValueSrc };
                        dtSrc.Rows.Add(drSrc);
                        break;
                    case DiffType.SheetMissingInSrc: // 比較先にのみシートが存在
                        drDst.ItemArray = new object[] { r.Id, r.FolderPathDst, r.FileNameDst, r.SheetName, r.CellAddress, r.CellValueDst };
                        dtDst.Rows.Add(drDst);
                        break;
                }
            }

            dgvSrc.DataSource = dtSrc;
            dgvDst.DataSource = dtDst;

            // ハイライト処理
            HighlightGrid(dgvSrc, resultsDict);
            HighlightGrid(dgvDst, resultsDict);

            dgvSrc.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgvDst.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }
        
        private void HighlightGrid(DataGridView dgv, Dictionary<int, DiffResult> results)
        {
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row.IsNewRow) continue;
                var id = int.Parse(row.Cells["ID"].Value.ToString());
                if (results.TryGetValue(id, out var result))
                {
                    switch (result.DiffType)
                    {
                        case DiffType.CellValueMismatch:
                            row.Cells["セル値"].Style.BackColor = System.Drawing.Color.Yellow;
                            break;
                        case DiffType.SheetMissingInDst:
                        case DiffType.SheetMissingInSrc:
                             row.Cells["シート名"].Style.BackColor = System.Drawing.Color.Yellow;
                            break;
                    }
                }
            }
        }
    }

    [DataContract]
    internal class SimpleExcelDiffSettings
    {
        [DataMember] public string PathSrc { get; set; }
        [DataMember] public string PathDst { get; set; }
        [DataMember] public bool EnableSubDir { get; set; }
    }
    
    // 差分種別をEnumで管理するように変更
    internal enum DiffType { CellValueMismatch, SheetMissingInDst, SheetMissingInSrc }

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
    }
}