using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace SimpleExcelBookSelector
{
    public partial class MainForm : Form
    {
        // Windows APIのインポート
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private Timer _timer;
        private List<string> _lastDisplayedKeys = new List<string>();

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
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            dataGridViewResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewResults.MultiSelect = false;
            dataGridViewResults.CellClick += DataGridViewResults_CellClick;

            RefreshExcelFileList();

            _timer = new Timer();
            _timer.Interval = 1000;
            _timer.Tick += (s, args) => RefreshExcelFileList();
            _timer.Start();
        }

        private void RefreshExcelFileList()
        {
            dynamic excelApp = null;
            var newIdentifiers = new List<ExcelSheetIdentifier>();
            var newKeys = new List<string>();

            try
            {
                excelApp = Marshal.GetActiveObject("Excel.Application");

                foreach (dynamic wb in excelApp.Workbooks)
                {
                    foreach (dynamic ws in wb.Sheets)
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
            catch (COMException) { }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }

            if (!_lastDisplayedKeys.SequenceEqual(newKeys))
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
                row.Tag = id;
            }

            dataGridViewResults.ResumeLayout();
        }

        private void DataGridViewResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dataGridViewResults.Rows[e.RowIndex];
            if (!(row.Tag is ExcelSheetIdentifier id)) return;

            dynamic excelApp = null;
            dynamic wb = null;
            dynamic ws = null;

            try
            {
                excelApp = Marshal.GetActiveObject("Excel.Application");
                
                // フルパスで目的のワークブックを検索（堅牢性を向上）
                foreach (dynamic book in excelApp.Workbooks)
                {
                    if (book.FullName == id.WorkbookFullName)
                    {
                        wb = book;
                        break;
                    }
                }

                if (wb != null)
                {
                    ws = wb.Sheets[id.WorksheetName];

                    // シートをアクティブ化
                    ws.Activate();

                    // Excelアプリケーションのウィンドウを最前面に表示
                    SetForegroundWindow((IntPtr)excelApp.Hwnd);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excelの操作に失敗しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
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
