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
        public List<string> FileHistory { get; private set; }
        private const string CheckBoxColumnName = "colCheck";
        private const string FilePathColumnName = "colFilePath";

        public HistoryForm(List<string> fileHistory)
        {
            InitializeComponent();
            FileHistory = new List<string>(fileHistory); // Create a copy
        }

        private void HistoryForm_Load(object sender, EventArgs e)
        {
            SetupDataGridView();
            PopulateDataGridView(FileHistory);
        }

        private void SetupDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoGenerateColumns = false;

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

        private void PopulateDataGridView(List<string> history)
        {
            dataGridView1.Rows.Clear();
            foreach (var path in history)
            {
                dataGridView1.Rows.Add(false, path);
            }
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
                PopulateDataGridView(FileHistory);
                textBox1.Clear();
                this.DialogResult = DialogResult.OK; // Notify MainForm to update
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
                FileHistory.RemoveAll(f => filesToRemove.Contains(f));
                textBox1_TextChanged(sender, e); // Re-apply filter
                this.DialogResult = DialogResult.OK; // Notify MainForm to update
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
            var filterText = textBox1.Text;

            if (string.IsNullOrWhiteSpace(filterText))
            {
                PopulateDataGridView(FileHistory);
            }
            else
            {
                var filteredHistory = FileHistory
                    .Where(path => path.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToList();
                PopulateDataGridView(filteredHistory);
            }
        }
    }
}