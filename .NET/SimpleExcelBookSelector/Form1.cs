using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;


namespace SimpleExcelBookSelector
{
    public partial class MainForm : Form
    {
        private const int MaxHistoryCount = 50;

        private AppSettings _settings;
        private readonly string _settingsFilePath = Path.Combine(
            Application.StartupPath,
            "settings.json");

        // Windows APIのインポート
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private Timer _timer;
        private List<string> _lastDisplayedKeys = new List<string>();
        private string _currentSortColumnName;
        private bool _isSortAscending = true;
        private static readonly StringComparer SortComparer = StringComparer.OrdinalIgnoreCase;


        private class ExcelSheetIdentifier
        {
            public string WorkbookFullName { get; set; }
            public string WorksheetName { get; set; }
            public string DirectoryName { get; set; }
            public string FileName { get; set; }
            public bool IsPinned { get; set; }
            public DateTime? LastUpdated { get; set; }
            public int OriginalIndex { get; set; }
            public string Key => $"{WorkbookFullName}|{WorksheetName}";
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
            dataGridViewResults.MultiSelect = true;
            dataGridViewResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;
            dataGridViewResults.CellClick += DataGridViewResults_CellClick;
            dataGridViewResults.CellDoubleClick += DataGridViewResults_CellDoubleClick;
            dataGridViewResults.CellToolTipTextNeeded += DataGridViewResults_CellToolTipTextNeeded;
            dataGridViewResults.ColumnHeaderMouseClick += DataGridViewResults_ColumnHeaderMouseClick;
            chkEnableSheetSelectMode.CheckedChanged += ChkEnableSheetSelectMode_CheckedChanged;
            chkEnableAutoUpdateMode.CheckedChanged += ChkEnableAutoUpdateMode_CheckedChanged;
            textAutoUpdateSec.TextChanged += TextAutoUpdateSec_TextChanged;
            btnForceUpdate.Click += BtnForceUpdate_Click;


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

        private void RefreshExcelFileList(bool forceUpdate = false)
        {
            dynamic excelApp = null;
            var newIdentifiers = new List<ExcelSheetIdentifier>();

            try
            {
                excelApp = Marshal.GetActiveObject("Excel.Application");

                int index = 0;
                foreach (dynamic wb in excelApp.Workbooks)
                {
                    AddToHistory(wb.FullName);

                    var historyItem = FindHistoryItem(wb.FullName);
                    bool isPinned = historyItem?.IsPinned ?? false;
                    DateTime? lastUpdated = historyItem?.LastUpdated;
                    string directoryName = Path.GetDirectoryName(wb.FullName);
                    string fileName = Path.GetFileName(wb.FullName);

                    if (chkEnableSheetSelectMode.Checked)
                    {
                        foreach (dynamic ws in wb.Sheets)
                        {
                            newIdentifiers.Add(new ExcelSheetIdentifier
                            {
                                WorkbookFullName = wb.FullName,
                                WorksheetName = ws.Name,
                                DirectoryName = directoryName,
                                FileName = fileName,
                                IsPinned = isPinned,
                                LastUpdated = lastUpdated,
                                OriginalIndex = index++
                            });
                        }
                    }
                    else
                    {
                        newIdentifiers.Add(new ExcelSheetIdentifier
                        {
                            WorkbookFullName = wb.FullName,
                            WorksheetName = string.Empty,
                            DirectoryName = directoryName,
                            FileName = fileName,
                            IsPinned = isPinned,
                            LastUpdated = lastUpdated,
                            OriginalIndex = index++
                        });
                    }
                }

            }
            catch (COMException) { }
            catch (ArgumentException) { }
            finally
            {
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            }

            var sortedIdentifiers = ApplyCurrentSort(newIdentifiers);
            var sortedKeys = sortedIdentifiers.Select(id => id.Key).ToList();

            if (forceUpdate || !_lastDisplayedKeys.SequenceEqual(sortedKeys))
            {
                UpdateDataGridView(sortedIdentifiers);
                _lastDisplayedKeys = sortedKeys;
            }
            else
            {
                UpdateSortGlyph();
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
                row.Cells["clmPinned"].Value = id.IsPinned ? "★" : string.Empty;
                row.Cells["clmDir"].Value = id.DirectoryName;
                row.Cells["clmFile"].Value = id.FileName;
                row.Cells["clmSheet"].Value = id.WorksheetName;
                row.Cells["clmUpdated"].Value = FormatTimestamp(id.LastUpdated);
                row.Tag = id;
            }

            dataGridViewResults.ResumeLayout();
            UpdateSortGlyph();
        }

        private List<ExcelSheetIdentifier> ApplyCurrentSort(List<ExcelSheetIdentifier> identifiers)
        {
            if (identifiers == null || identifiers.Count == 0)
            {
                return new List<ExcelSheetIdentifier>();
            }

            IOrderedEnumerable<ExcelSheetIdentifier> ordered;

            switch (_currentSortColumnName)
            {
                case "clmPinned":
                    ordered = _isSortAscending
                        ? identifiers.OrderBy(id => id.IsPinned)
                        : identifiers.OrderByDescending(id => id.IsPinned);
                    break;
                case "clmDir":
                    ordered = _isSortAscending
                        ? identifiers.OrderBy(id => id.DirectoryName, SortComparer)
                        : identifiers.OrderByDescending(id => id.DirectoryName, SortComparer);
                    break;
                case "clmFile":
                    ordered = _isSortAscending
                        ? identifiers.OrderBy(id => id.FileName, SortComparer)
                        : identifiers.OrderByDescending(id => id.FileName, SortComparer);
                    break;
                case "clmSheet":
                    ordered = _isSortAscending
                        ? identifiers.OrderBy(id => id.WorksheetName, SortComparer)
                        : identifiers.OrderByDescending(id => id.WorksheetName, SortComparer);
                    break;
                case "clmUpdated":
                    ordered = _isSortAscending
                        ? identifiers.OrderBy(id => id.LastUpdated ?? DateTime.MinValue)
                        : identifiers.OrderByDescending(id => id.LastUpdated ?? DateTime.MinValue);
                    break;
                default:
                    ordered = identifiers.OrderBy(id => id.OriginalIndex);
                    break;
            }

            ordered = ordered
                .ThenBy(id => id.DirectoryName, SortComparer)
                .ThenBy(id => id.FileName, SortComparer)
                .ThenBy(id => id.WorksheetName, SortComparer)
                .ThenBy(id => id.OriginalIndex);

            return ordered.ToList();
        }

        private void UpdateSortGlyph()
        {
            foreach (DataGridViewColumn column in dataGridViewResults.Columns)
            {
                if (column.Name == _currentSortColumnName)
                {
                    column.HeaderCell.SortGlyphDirection = _isSortAscending ? SortOrder.Ascending : SortOrder.Descending;
                }
                else
                {
                    column.HeaderCell.SortGlyphDirection = SortOrder.None;
                }
            }
        }

        private static string FormatTimestamp(DateTime? value)
        {
            return value.HasValue ? value.Value.ToString("yyyy/MM/dd HH:mm:ss") : string.Empty;
        }

        private HistoryItem FindHistoryItem(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || _settings?.FileHistory == null)
            {
                return null;
            }

            return _settings.FileHistory.FirstOrDefault(item => item.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
        }

        private bool IsPinned(string filePath)
        {
            var item = FindHistoryItem(filePath);
            return item?.IsPinned ?? false;
        }

        // For migration from old settings format
        [DataContract]
        private class AppSettings_Old
        {
            [DataMember]
            public bool IsSheetSelectionEnabled { get; set; }
            [DataMember]
            public bool IsAutoRefreshEnabled { get; set; }
            [DataMember]
            public int RefreshInterval { get; set; }
            [DataMember]
            public List<string> FileHistory { get; set; }
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsFilePath))
            {
                _settings = new AppSettings();
                return;
            }

            try
            {
                // Try to load the new format first
                using (var stream = File.OpenRead(_settingsFilePath))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    _settings = (AppSettings)serializer.ReadObject(stream);
                }
            }
            catch (SerializationException)
            {
                // If it fails, try to load the old format and migrate
                try
                {
                    using (var stream = File.OpenRead(_settingsFilePath))
                    {
                        var serializer = new DataContractJsonSerializer(typeof(AppSettings_Old));
                        var oldSettings = (AppSettings_Old)serializer.ReadObject(stream);

                        _settings = new AppSettings
                        {
                            IsSheetSelectionEnabled = oldSettings.IsSheetSelectionEnabled,
                            IsAutoRefreshEnabled = oldSettings.IsAutoRefreshEnabled,
                            RefreshInterval = oldSettings.RefreshInterval,
                            FileHistory = oldSettings.FileHistory.Select(path => new HistoryItem { FilePath = path, IsPinned = false, LastUpdated = DateTime.Now }).ToList()
                        };
                    }
                    // Immediately save the settings in the new format
                    // SaveSettings(); // BUG: This overwrites migrated settings with default UI values.
                }
                catch
                {
                    // If migration also fails, create default settings
                    _settings = new AppSettings();
                }
            }
            catch
            {
                // For any other error, create default settings
                _settings = new AppSettings();
            }

            // Reflect settings on the UI
            chkEnableSheetSelectMode.Checked = _settings.IsSheetSelectionEnabled;
            chkEnableAutoUpdateMode.Checked = _settings.IsAutoRefreshEnabled;
            textAutoUpdateSec.Text = _settings.RefreshInterval.ToString();
            chkIsOpenDir.Checked = _settings.IsOpenFolderOnDoubleClickEnabled;

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
            _settings.IsOpenFolderOnDoubleClickEnabled = chkIsOpenDir.Checked;

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
                MessageBox.Show($"Failed to save settings.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddToHistory(string filePath)
        {
            if (string.IsNullOrEmpty(filePath)) return;

            DateTime? lastModified = null;
            try
            {
                if (File.Exists(filePath))
                {
                    lastModified = File.GetLastWriteTime(filePath);
                }
            }
            catch
            {
                lastModified = DateTime.Now;
            }

            var existingItem = _settings.FileHistory.FirstOrDefault(p => p.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));

            if (existingItem != null)
            {
                // If item exists, move it to the top of its group (pinned or not pinned)
                _settings.FileHistory.Remove(existingItem);
                if (lastModified.HasValue)
                {
                    existingItem.LastUpdated = lastModified;
                }
                int insertIndex = _settings.FileHistory.TakeWhile(i => i.IsPinned).Count();
                _settings.FileHistory.Insert(existingItem.IsPinned ? 0 : insertIndex, existingItem);
            }
            else
            {
                // If new, add as not pinned
                var newItem = new HistoryItem { FilePath = filePath, IsPinned = false, LastUpdated = lastModified ?? DateTime.Now };
                int insertIndex = _settings.FileHistory.TakeWhile(i => i.IsPinned).Count();
                _settings.FileHistory.Insert(insertIndex, newItem);
            }

            // Trim non-pinned items if history exceeds max count
            var pinnedCount = _settings.FileHistory.Count(i => i.IsPinned);
            var notPinnedCount = _settings.FileHistory.Count(i => !i.IsPinned);

            if (pinnedCount + notPinnedCount > MaxHistoryCount)
            {
                var itemsToRemove = _settings.FileHistory
                    .Where(i => !i.IsPinned)
                    .Skip(MaxHistoryCount - pinnedCount)
                    .ToList();

                foreach (var item in itemsToRemove)
                {
                    _settings.FileHistory.Remove(item);
                }
            }
        }

        private void BtnForceUpdate_Click(object sender, EventArgs e)
        {
            RefreshExcelFileList(forceUpdate: true);
        }

        private void DataGridViewResults_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0)
            {
                return;
            }

            var column = dataGridViewResults.Columns[e.ColumnIndex];
            if (column == null)
            {
                return;
            }

            var columnName = column.Name;

            if (_currentSortColumnName == columnName)
            {
                _isSortAscending = !_isSortAscending;
            }
            else
            {
                _currentSortColumnName = columnName;
                _isSortAscending = true;
            }

            RefreshExcelFileList(forceUpdate: true);
        }

        private void DataGridViewResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex == -1)
            {
                dataGridViewResults.SelectAll();
                return;
            }

            if (e.RowIndex < 0) return;

            ActivateExcelForRowIndex(e.RowIndex);
        }

        private void DataGridViewResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var row = dataGridViewResults.Rows[e.RowIndex];
            if (!(row.Tag is ExcelSheetIdentifier id)) return;

            try
            {
                if (chkIsOpenDir.Checked)
                {
                    string folderPath = Path.GetDirectoryName(id.WorkbookFullName);
                    if (Directory.Exists(folderPath))
                    {
                        Process.Start(folderPath);
                    }
                }
                else
                {
                    if (File.Exists(id.WorkbookFullName))
                    {
                        Process.Start(id.WorkbookFullName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open folder or file.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                textAutoUpdateSec.TextChanged -= TextAutoUpdateSec_TextChanged;
                textAutoUpdateSec.Text = "1";
                _timer.Interval = 1000;
                textAutoUpdateSec.TextChanged += TextAutoUpdateSec_TextChanged;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
            _timer?.Stop();
            _timer?.Dispose();
        }

        private void DataGridViewResults_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > -1)
            {
                var cell = dataGridViewResults.Rows[e.RowIndex].Cells[e.ColumnIndex];
                e.ToolTipText = cell.Value?.ToString();
            }
        }

        private void btnHistory_Click(object sender, EventArgs e)
        {
            _settings.IsOpenFolderOnDoubleClickEnabled = chkIsOpenDir.Checked;
            using (var historyForm = new HistoryForm(_settings))
            {
                if (historyForm.ShowDialog(this) == DialogResult.OK)
                {
                    // History was changed in the dialog, perform a deep copy of the results
                    _settings.FileHistory = new List<HistoryItem>(historyForm.FileHistory.Select(i => new HistoryItem { FilePath = i.FilePath, IsPinned = i.IsPinned, LastUpdated = i.LastUpdated }));
                    SaveSettings();
                    RefreshExcelFileList(forceUpdate: true);
                }
            }
        }

        private void ActivateExcelForRowIndex(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= dataGridViewResults.Rows.Count)
            {
                return;
            }

            var row = dataGridViewResults.Rows[rowIndex];
            if (!(row.Tag is ExcelSheetIdentifier id))
            {
                return;
            }

            ActivateExcelEntry(id);
        }

        private void ActivateExcelEntry(ExcelSheetIdentifier id)
        {
            if (id == null)
            {
                return;
            }

            dynamic excelApp = null;
            dynamic wb = null;
            dynamic ws = null;

            try
            {
                excelApp = Marshal.GetActiveObject("Excel.Application");

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
                    if (chkEnableSheetSelectMode.Checked && !string.IsNullOrEmpty(id.WorksheetName))
                    {
                        ws = wb.Sheets[id.WorksheetName];
                        ws.Activate();
                    }
                    else
                    {
                        wb.Activate();
                    }

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

            BringMainFormToFront();
        }

        private void BringMainFormToFront()
        {
            if (IsDisposed)
            {
                return;
            }

            BeginInvoke(new Action(() =>
            {
                if (IsDisposed)
                {
                    return;
                }

                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Normal;
                }

                bool originalTopMost = TopMost;
                if (!originalTopMost)
                {
                    TopMost = true;
                }

                Activate();
                SetForegroundWindow(Handle);

                if (!originalTopMost)
                {
                    TopMost = false;
                }
            }));
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if ((keyData & Keys.Control) == Keys.Control && (keyData & Keys.KeyCode) == Keys.Tab)
            {
                if (!(dataGridViewResults.Focused || dataGridViewResults.ContainsFocus))
                {
                    return base.ProcessCmdKey(ref msg, keyData);
                }

                bool isReverse = (keyData & Keys.Shift) == Keys.Shift;
                MoveSelectionByCtrlTab(isReverse ? -1 : 1);
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void MoveSelectionByCtrlTab(int direction)
        {
            if (dataGridViewResults.Rows.Count == 0)
            {
                return;
            }

            int rowCount = dataGridViewResults.Rows.Count;
            int currentIndex = dataGridViewResults.CurrentCell?.RowIndex ?? -1;

            if (currentIndex < 0 && dataGridViewResults.SelectedRows.Count > 0)
            {
                currentIndex = dataGridViewResults.SelectedRows[0].Index;
            }

            if (currentIndex < 0)
            {
                currentIndex = direction > 0 ? -1 : rowCount;
            }

            int targetIndex = direction > 0
                ? (currentIndex + 1 + rowCount) % rowCount
                : (currentIndex - 1 + rowCount) % rowCount;

            dataGridViewResults.CurrentCell = dataGridViewResults.Rows[targetIndex].Cells[0];
            dataGridViewResults.ClearSelection();
            dataGridViewResults.Rows[targetIndex].Selected = true;

            try
            {
                if (targetIndex >= 0 && targetIndex < dataGridViewResults.RowCount)
                {
                    dataGridViewResults.FirstDisplayedScrollingRowIndex = targetIndex;
                }
            }
            catch (InvalidOperationException)
            {
                // 行をスクロールできない場合は無視
            }

            ActivateExcelForRowIndex(targetIndex);
        }
    }
}
