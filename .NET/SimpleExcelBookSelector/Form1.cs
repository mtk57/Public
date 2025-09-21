using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SimpleExcelBookSelector
{
    public partial class MainForm : Form
    {
        private Timer _timer;
        private List<string> _lastDisplayedKeys = new List<string>();

        // DataGridViewの行に格納する識別情報用のクラス
        private class ExcelSheetIdentifier
        {
            public string WorkbookFullName { get; set; }
            public string WorksheetName { get; set; }
        }

        public MainForm()
        {
            InitializeComponent();
            this.Load += MainForm_Load;
            this.FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            dataGridViewResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewResults.MultiSelect = false;
            dataGridViewResults.CellClick += DataGridViewResults_CellClick;

            RefreshExcelFileList();

            _timer = new Timer();
            _timer.Interval = 1000; // 1秒
            _timer.Tick += (s, args) => RefreshExcelFileList();
            _timer.Start();
        }

        private void RefreshExcelFileList()
        {
            Excel.Application excelApp = null;
            var newIdentifiers = new List<ExcelSheetIdentifier>();
            var newKeys = new List<string>();

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    foreach (Excel.Worksheet ws in wb.Sheets)
                    {
                        newKeys.Add($"{wb.FullName}|{ws.Name}");
                        newIdentifiers.Add(new ExcelSheetIdentifier
                        {
                            WorkbookFullName = wb.FullName,
                            WorksheetName = ws.Name
                        });
                    }
                }
            }
            catch (COMException)
            {
                // Excelが実行されていない場合
            }
            finally
            {
                // このメソッドで取得したCOMオブジェクトはここで解放
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }

            bool hasChanges = !_lastDisplayedKeys.SequenceEqual(newKeys);

            if (hasChanges)
            {
                UpdateDataGridView(newIdentifiers);
                _lastDisplayedKeys = newKeys;
            }
        }

        private void UpdateDataGridView(List<ExcelSheetIdentifier> identifiers)
        {
            dataGridViewResults.SuspendLayout();
            dataGridViewResults.Rows.Clear();

            foreach (var id in identifiers)
            {
                var rowIndex = dataGridViewResults.Rows.Add();
                var row = dataGridViewResults.Rows[rowIndex];
                row.Cells["clmDir"].Value = Path.GetDirectoryName(id.WorkbookFullName);
                row.Cells["clmFile"].Value = Path.GetFileName(id.WorkbookFullName);
                row.Cells["clmSheet"].Value = id.WorksheetName;
                row.Tag = id; // 行のTagプロパティに識別情報を格納
            }

            dataGridViewResults.ResumeLayout();
        }

        private void DataGridViewResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dataGridViewResults.Rows[e.RowIndex];
            if (!(row.Tag is ExcelSheetIdentifier id)) return;

            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                // クリック時に改めてExcelオブジェクトを取得
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                wb = excelApp.Workbooks[Path.GetFileName(id.WorkbookFullName)];
                ws = wb.Sheets[id.WorksheetName];

                // ウィンドウをアクティブ化
                excelApp.Visible = true;
                ((Excel.Window)wb.Windows[1]).Activate();
                ws.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excelの操作に失敗しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // このメソッドで取得したCOMオブジェクトをすべて解放
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _timer?.Stop();
            _timer?.Dispose();
        }
    }
}