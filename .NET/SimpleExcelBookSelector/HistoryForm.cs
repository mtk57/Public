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
        private const string PinnedColumnName = "colPinned";
        private const string CheckBoxColumnName = "colCheck";
        private const string FilePathColumnName = "colFilePath";

        public HistoryForm(List<HistoryItem> fileHistory)
        {
            InitializeComponent();
            FileHistory = new List<HistoryItem>(fileHistory.Select(item => new HistoryItem { FilePath = item.FilePath, IsPinned = item.IsPinned }));
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
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter }
            };
            dataGridView1.Columns.Add(pinnedColumn);

            var checkBoxColumn = new DataGridViewCheckBoxColumn
            {
                Name = CheckBoxColumnName,
                HeaderText = "",
                Width = 30
            };
            dataGridView1.Columns.Add(checkBoxColumn);

            var filePathColumn = new DataGridViewTextBoxColumn
            {
                Name = FilePathColumnName,
                HeaderText = "ファイルパス",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };
            dataGridView1.Columns.Add(filePathColumn);
        }

        private void PopulateDataGridView(List<HistoryItem> history)
        {
            dataGridView1.Rows.Clear();
            foreach (var item in history)
            {
                dataGridView1.Rows.Add(item.IsPinned ? "★" : "", false, item.FilePath);
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

            var sortedHistory = filteredHistory.OrderByDescending(item => item.IsPinned).ToList();
            
            PopulateDataGridView(sortedHistory);
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

        private void btnAllClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("全ての履歴を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                FileHistory.Clear();
                ApplyFilterAndSort();
                this.DialogResult = DialogResult.OK;
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
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[CheckBoxColumnName].Value = false;
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

            if (filesToRemove.Count == 0)
            {
                MessageBox.Show("削除するファイルが選択されていません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (MessageBox.Show($"{filesToRemove.Count}件の履歴を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
            {
                FileHistory.RemoveAll(item => filesToRemove.Contains(item.FilePath));
                ApplyFilterAndSort();
                this.DialogResult = DialogResult.OK;
            }
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var filePath = dataGridView1.Rows[e.RowIndex].Cells[FilePathColumnName].Value.ToString();
                OpenFiles(new[] { filePath });
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        private void btnClearFilter_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            // The TextChanged event will handle the update
        }

        private void btnPinnedSelectedFiles_Click(object sender, EventArgs e)
        {
            var selectedPaths = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells[CheckBoxColumnName].Value))
                {
                    selectedPaths.Add(row.Cells[FilePathColumnName].Value.ToString());
                }
            }

            if (selectedPaths.Count == 0)
            {
                MessageBox.Show("ピン留め/解除するファイルが選択されていません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var path in selectedPaths)
            {
                var item = FileHistory.FirstOrDefault(i => i.FilePath == path);
                if (item != null)
                {
                    item.IsPinned = !item.IsPinned; // Toggle pinned state
                }
            }

            ApplyFilterAndSort();
            this.DialogResult = DialogResult.OK;
        }
    }
}
