using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;
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
            LoadSettings(); // Load settings first
            dataGridViewResults.Columns["clmSheet"].Visible = _settings.IsSheetSelectionEnabled;

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            dataGridViewResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewResults.MultiSelect = false;
            dataGridViewResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;
            dataGridViewResults.CellClick += DataGridViewResults_CellClick;
            dataGridViewResults.CellDoubleClick += DataGridViewResults_CellDoubleClick;
            chkEnableSheetSelectMode.CheckedChanged += ChkEnableSheetSelectMode_CheckedChanged;
            chkEnableAutoUpdateMode.CheckedChanged += ChkEnableAutoUpdateMode_CheckedChanged;
            textAutoUpdateSec.TextChanged += TextAutoUpdateSec_TextChanged;


            RefreshExcelFileList();

            _timer = new Timer();
            _timer.Tick += (s, args) => RefreshExcelFileList();
            // Apply settings for the timer
            textAutoUpdateSec.Enabled = _settings.IsAutoRefreshEnabled;
            TextAutoUpdateSec_TextChanged(null, null); // Set interval and validate

            if (_settings.IsAutoRefreshEnabled)
            {
                _timer.Start();
            }
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
                    if (chkEnableSheetSelectMode.Checked)
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
                    else
                    {
                        newKeys.Add(wb.FullName);
                        newIdentifiers.Add(new ExcelSheetIdentifier
                        {
                            WorkbookFullName = wb.FullName,
                            WorksheetName = string.Empty
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

        private AppSettings _settings;
        private readonly string _settingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SimpleExcelBookSelector",
            "settings.json");

        private void LoadSettings()
        {
            if (File.Exists(_settingsFilePath))
            {
                try
                {
                    using (var stream = File.OpenRead(_settingsFilePath))
                    {
                        var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                        _settings = (AppSettings)serializer.ReadObject(stream);
                    }
                }
                catch
                {
                    // If the file is corrupted, overwrite with default settings
                    _settings = new AppSettings();
                }
            }
            else
            {
                _settings = new AppSettings();
            }

            // Reflect settings on the UI
            chkEnableSheetSelectMode.Checked = _settings.IsSheetSelectionEnabled;
            chkEnableAutoUpdateMode.Checked = _settings.IsAutoRefreshEnabled;
            textAutoUpdateSec.Text = _settings.RefreshInterval.ToString();
        }

        private void SaveSettings()
        {
            // Get settings from the UI
            _settings.IsSheetSelectionEnabled = chkEnableSheetSelectMode.Checked;
            _settings.IsAutoRefreshEnabled = chkEnableAutoUpdateMode.Checked;
            if (int.TryParse(textAutoUpdateSec.Text, out int interval) && interval > 0)
            {
                _settings.RefreshInterval = interval;
            }
            else
            {
                _settings.RefreshInterval = 1; // Use default for invalid values
            }

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_settingsFilePath));
                using (var stream = File.Create(_settingsFilePath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    serializer.WriteObject(stream, _settings);
                }
            }
            catch (Exception ex)
            {
                // Error handling (e.g., notify with a MessageBox)
                MessageBox.Show($"Failed to save settings.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DataGridViewResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Handle top-left header cell click for select-all
            if (e.RowIndex == -1 && e.ColumnIndex == -1)
            {
                dataGridViewResults.SelectAll();
                return;
            }

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
                    if (chkEnableSheetSelectMode.Checked)
                    {
                        ws = wb.Sheets[id.WorksheetName];

                        // Activate the sheet
                        ws.Activate();
                    }
                    else
                    {
                        // Activate the workbook without changing the sheet
                        wb.Activate();
                    }

                    // Bring the Excel application window to the foreground
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

        private void DataGridViewResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dataGridViewResults.Rows[e.RowIndex];
            if (!(row.Tag is ExcelSheetIdentifier id)) return;

            try
            {
                string folderPath = Path.GetDirectoryName(id.WorkbookFullName);
                if (Directory.Exists(folderPath))
                {
                    Process.Start(folderPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open folder.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChkEnableSheetSelectMode_CheckedChanged(object sender, EventArgs e)
        {
            dataGridViewResults.Columns["clmSheet"].Visible = chkEnableSheetSelectMode.Checked;
            RefreshExcelFileList();
        }

        private void ChkEnableAutoUpdateMode_CheckedChanged(object sender, EventArgs e)
        {
            textAutoUpdateSec.Enabled = chkEnableAutoUpdateMode.Checked;
            if (chkEnableAutoUpdateMode.Checked)
            {
                _timer.Start();
            }
            else
            {
                _timer.Stop();
            }
        }

        private void TextAutoUpdateSec_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(textAutoUpdateSec.Text, out int interval) && interval > 0)
            {
                _timer.Interval = interval * 1000;
            }
            else
            {
                // A handler is attached to TextChanged, so changing the text here will cause a recursive call.
                // To avoid this, we detach the handler, change the text, and then reattach it.
                textAutoUpdateSec.TextChanged -= TextAutoUpdateSec_TextChanged;
                textAutoUpdateSec.Text = "1";
                _timer.Interval = 1000;
                textAutoUpdateSec.TextChanged += TextAutoUpdateSec_TextChanged;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings(); // Save settings

            _timer?.Stop();
            _timer?.Dispose();
        }
    }
}
