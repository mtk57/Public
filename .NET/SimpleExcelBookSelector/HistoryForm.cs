using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SimpleExcelBookSelector
{
    public partial class HistoryForm : Form
    {
        public List<HistoryItem> FileHistory { get; private set; }
        private readonly AppSettings _settings;
        private const string PinnedColumnName = "colPinned";
        private const string CheckBoxColumnName = "colCheck";
        private const string FilePathColumnName = "colFilePath";
        private string _currentSortColumnName = PinnedColumnName;
        private bool _isSortAscending = false;
        private static readonly StringComparer SortComparer = StringComparer.OrdinalIgnoreCase;

        public HistoryForm(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            FileHistory = new List<HistoryItem>(settings.FileHistory.Select(item => new HistoryItem { FilePath = item.FilePath, IsPinned = item.IsPinned }));
        }

        private void HistoryForm_Load(object sender, EventArgs e)
        {
            SetupDataGridView();
            ApplyFilterAndSort();
        }

        private void SetupDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoGenerateColumns = false;

            var pinnedColumn = new DataGridViewTextBoxColumn
            {
                Name = PinnedColumnName,
                HeaderText = "ピン",
                Width = 35,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter },
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            dataGridView1.Columns.Add(pinnedColumn);

            var checkBoxColumn = new DataGridViewCheckBoxColumn
            {
                Name = CheckBoxColumnName,
                HeaderText = "",
                Width = 30,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dataGridView1.Columns.Add(checkBoxColumn);

            var filePathColumn = new DataGridViewTextBoxColumn
            {
                Name = FilePathColumnName,
                HeaderText = "ファイルパス",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                SortMode = DataGridViewColumnSortMode.Programmatic
            };
            dataGridView1.Columns.Add(filePathColumn);

            dataGridView1.CellClick += DataGridView1_CellClick;
        }

        private void PopulateDataGridView(List<HistoryItem> history)
        {
            dataGridView1.Rows.Clear();
            foreach (var item in history)
            {
                dataGridView1.Rows.Add(item.IsPinned ? "★" : "", false, item.FilePath);
            }
        }

        private List<HistoryItem> SortHistory(IEnumerable<HistoryItem> history)
        {
            IOrderedEnumerable<HistoryItem> ordered;

            switch (_currentSortColumnName)
            {
                case PinnedColumnName:
                    ordered = (_isSortAscending
                        ? history.OrderBy(item => item.IsPinned)
                        : history.OrderByDescending(item => item.IsPinned))
                        .ThenBy(item => item.FilePath, SortComparer);
                    break;
                case FilePathColumnName:
                    ordered = _isSortAscending
                        ? history.OrderBy(item => item.FilePath, SortComparer)
                        : history.OrderByDescending(item => item.FilePath, SortComparer);
                    ordered = ordered.ThenByDescending(item => item.IsPinned);
                    break;
                default:
                    ordered = history.OrderByDescending(item => item.IsPinned)
                        .ThenBy(item => item.FilePath, SortComparer);
                    break;
            }

            return ordered.ToList();
        }

        private void UpdateSortGlyphs()
        {
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (column.SortMode == DataGridViewColumnSortMode.NotSortable)
                {
                    column.HeaderCell.SortGlyphDirection = SortOrder.None;
                    continue;
                }

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

        private void ApplyFilterAndSort()
        {
            var filterText = textBox1.Text;

            IEnumerable<HistoryItem> filteredHistory = FileHistory;

            if (!string.IsNullOrWhiteSpace(filterText))
            {
                filteredHistory = filteredHistory
                    .Where(item => item.FilePath.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0);
            }

            var sortedHistory = SortHistory(filteredHistory);

            PopulateDataGridView(sortedHistory);
            UpdateSortGlyphs();
        }

        private void btnAllOpen_Click(object sender, EventArgs e)
        {
            var filesToOpen = dataGridView1.Rows
                .Cast<DataGridViewRow>()
                .Select(row => row.Cells[FilePathColumnName].Value.ToString())
                .ToList();

            OpenFiles(filesToOpen);
        }

        private void btnSelectOpen_Click(object sender, EventArgs e)
        {
            var selectedFiles = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[CheckBoxColumnName].Value))
                {
                    selectedFiles.Add(row.Cells[FilePathColumnName].Value.ToString());
                }
            }
            OpenFiles(selectedFiles);
        }


        private void btnOpenAllPin_Click(object sender, EventArgs e)
        {
            var pinnedPaths = FileHistory
                .Where(item => item.IsPinned)
                .Select(item => item.FilePath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (pinnedPaths.Count == 0)
            {
                MessageBox.Show("ピン留めされた履歴がありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            OpenFiles(pinnedPaths);
        }

        private void btnAllClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("全ての履歴を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                FileHistory.Clear();
                ApplyFilterAndSort();
            }
        }

        private void OpenFiles(IEnumerable<string> filePaths)
        {
            foreach (var path in filePaths)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show($"ファイルが見つかりません: {path}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイルを開けませんでした: {path}\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnUnselectedAll_Click(object sender, EventArgs e)
        {
            bool shouldCheckAll = dataGridView1.Rows
                .Cast<DataGridViewRow>()
                .Any(row => !(row.Cells[CheckBoxColumnName].Value is bool checkedValue && checkedValue));

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[CheckBoxColumnName].Value = shouldCheckAll;
            }
        }

        private void btnDeleteSelectedFiles_Click(object sender, EventArgs e)
        {
            var filesToRemove = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[CheckBoxColumnName].Value))
                {
                    filesToRemove.Add(row.Cells[FilePathColumnName].Value.ToString());
                }
            }

            if (filesToRemove.Count == 0) return;

            if (MessageBox.Show($"{filesToRemove.Count}件の履歴を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                FileHistory.RemoveAll(item => filesToRemove.Contains(item.FilePath));
                ApplyFilterAndSort();
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            var filePath = dataGridView1.Rows[e.RowIndex].Cells[FilePathColumnName].Value.ToString();

            try
            {
                if (_settings.IsOpenFolderOnDoubleClickEnabled)
                {
                    string folderPath = Path.GetDirectoryName(filePath);
                    if (Directory.Exists(folderPath))
                    {
                        Process.Start(folderPath);
                    }
                }
                else
                {
                    if (File.Exists(filePath))
                    {
                        Process.Start(filePath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open folder or file.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0)
            {
                return;
            }

            var column = dataGridView1.Columns[e.ColumnIndex];
            if (column == null || column.SortMode == DataGridViewColumnSortMode.NotSortable)
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

            ApplyFilterAndSort();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        private void btnClearFilter_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
        }

        private void ChangePinnedState(bool shouldBePinned)
        {
            var selectedPaths = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[CheckBoxColumnName].Value))
                {
                    selectedPaths.Add(row.Cells[FilePathColumnName].Value.ToString());
                }
            }

            if (selectedPaths.Count == 0) return;

            foreach (var path in selectedPaths)
            {
                var item = FileHistory.FirstOrDefault(i => i.FilePath == path);
                if (item != null)
                {
                    item.IsPinned = shouldBePinned;
                }
            }

            ApplyFilterAndSort();
        }

        private void btnPinnedSelectedFiles_Click(object sender, EventArgs e)
        {
            ChangePinnedState(true); // Pin
        }

        private void btnUnPinnedSelectedFiles_Click(object sender, EventArgs e)
        {
            ChangePinnedState(false); // Unpin
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            if (!string.Equals(dataGridView1.Columns[e.ColumnIndex].Name, PinnedColumnName, StringComparison.Ordinal))
            {
                return;
            }

            var filePath = dataGridView1.Rows[e.RowIndex].Cells[FilePathColumnName].Value?.ToString();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            var historyItem = FileHistory.FirstOrDefault(i => string.Equals(i.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
            if (historyItem == null)
            {
                return;
            }

            historyItem.IsPinned = !historyItem.IsPinned;
            ApplyFilterAndSort();
        }
    }
}
