using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
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
        private ToolTip _toolTip;


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

            _toolTip = new ToolTip();
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
            chkEnableSheetSelectMode.CheckedChanged += ChkEnableSheetSelectMode_CheckedChanged;
            chkEnableAutoUpdateMode.CheckedChanged += ChkEnableAutoUpdateMode_CheckedChanged;
            textAutoUpdateSec.TextChanged += TextAutoUpdateSec_TextChanged;
            cmbHistory.DrawMode = DrawMode.OwnerDrawFixed;
            cmbHistory.DrawItem += CmbHistory_DrawItem;
            cmbHistory.DropDownClosed += CmbHistory_DropDownClosed;
            cmbHistory.SelectedIndexChanged += CmbHistory_SelectedIndexChanged;
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
                    AddToHistory(wb.FullName);

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

                UpdateHistoryComboBox();
            }
            catch (COMException) { }
            catch (ArgumentException) { }
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
                            FileHistory = oldSettings.FileHistory.Select(path => new HistoryItem { FilePath = path, IsPinned = false }).ToList()
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

            UpdateHistoryComboBox();
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

            var existingItem = _settings.FileHistory.FirstOrDefault(p => p.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));

            if (existingItem != null)
            {
                // If item exists, move it to the top of its group (pinned or not pinned)
                _settings.FileHistory.Remove(existingItem);
                int insertIndex = _settings.FileHistory.TakeWhile(i => i.IsPinned).Count();
                _settings.FileHistory.Insert(existingItem.IsPinned ? 0 : insertIndex, existingItem);
            }
            else
            {
                // If new, add as not pinned
                var newItem = new HistoryItem { FilePath = filePath, IsPinned = false };
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

        private void UpdateHistoryComboBox()
        {
            cmbHistory.BeginUpdate();
            cmbHistory.Items.Clear();
            // Sort by Pinned status then by original order (which is most recent)
            var sortedHistory = _settings.FileHistory
                .OrderByDescending(i => i.IsPinned)
                .Select(i => i.FilePath)
                .ToArray();

            cmbHistory.Items.AddRange(sortedHistory);
            cmbHistory.EndUpdate();
        }

        private void CmbHistory_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbHistory.SelectedIndex == -1 || cmbHistory.SelectedItem == null) return;

            _timer?.Stop(); // タイマーを一時停止

            string filePath = cmbHistory.SelectedItem.ToString();

            try
            {
                if (File.Exists(filePath))
                {
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // 履歴から見つからないファイルを削除
                    _settings.FileHistory.RemoveAll(item => item.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
                    UpdateHistoryComboBox();
                    SaveSettings();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルを開けませんでした。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 自動更新が有効な場合のみタイマーを再開
                if (chkEnableAutoUpdateMode.Checked)
                {
                    _timer?.Start();
                }
            }
        }

        private void BtnForceUpdate_Click(object sender, EventArgs e)
        {
            RefreshExcelFileList();
        }

        private void DataGridViewResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
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

        private void CmbHistory_DrawItem(object sender, DrawItemEventArgs e)
        {if (e.Index < 0) return;

            e.DrawBackground();

            TextRenderer.DrawText(e.Graphics, cmbHistory.Items[e.Index].ToString(), e.Font, e.Bounds, e.ForeColor, TextFormatFlags.Left);

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                _toolTip.Show(cmbHistory.Items[e.Index].ToString(), cmbHistory, e.Bounds.Right, e.Bounds.Bottom);
            }

            e.DrawFocusRectangle();

        }

        private void CmbHistory_DropDownClosed(object sender, EventArgs e)
        {
            _toolTip.Hide(cmbHistory);
        }

        private void btnHistory_Click(object sender, EventArgs e)
        {
            _settings.IsOpenFolderOnDoubleClickEnabled = chkIsOpenDir.Checked;
            using (var historyForm = new HistoryForm(_settings))
            {
                if (historyForm.ShowDialog(this) == DialogResult.OK)
                {
                    // History was changed in the dialog, perform a deep copy of the results
                    _settings.FileHistory = new List<HistoryItem>(historyForm.FileHistory.Select(i => new HistoryItem { FilePath = i.FilePath, IsPinned = i.IsPinned }));
                    UpdateHistoryComboBox();
                    SaveSettings();
                }
            }
        }
    }
}