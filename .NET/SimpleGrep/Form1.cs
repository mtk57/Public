using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            this.button1.Click += new System.EventHandler(this.btnGrep_Click);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    cmbFolderPath.Text = fbd.SelectedPath;
                }
            }
        }

        private async void btnGrep_Click(object sender, EventArgs e)
        {
            string folderPath = cmbFolderPath.Text;
            string filePattern = comboBox1.Text;
            string grepPattern = cmbKeyword.Text;

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath) || string.IsNullOrWhiteSpace(filePattern) || string.IsNullOrWhiteSpace(grepPattern))
            {
                MessageBox.Show("検索フォルダー、ファイルパターン、検索パターンを正しく入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            dataGridViewResults.Rows.Clear();
            button1.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            try
            {
                await Task.Run(() =>
                {
                    SearchFiles(folderPath, filePattern, grepPattern);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button1.Enabled = true;
                this.Cursor = Cursors.Default;
                MessageBox.Show("検索が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SearchFiles(string folderPath, string filePattern, string grepPattern)
        {
            try
            {
                var searchOption = chkSearchSubDir.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                foreach (var filePath in Directory.EnumerateFiles(folderPath, filePattern, searchOption))
                {
                    try
                    {
                        using (var reader = new StreamReader(filePath))
                        {
                            string line;
                            int lineNumber = 1;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (line.Contains(grepPattern))
                                {
                                    this.Invoke((Action)(() =>
                                    {
                                        dataGridViewResults.Rows.Add(filePath, lineNumber, line);
                                    }));
                                }
                                lineNumber++;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // ファイル読み取りエラーはスキップ
                    }
                }
            }
            catch (Exception ex)
            {
                this.Invoke((Action)(() =>
                {
                    MessageBox.Show($"ディレクトリの検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
        }

        private void cmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            { 
                e.Effect = DragDropEffects.None;
            }
        }

        private void cmbFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length > 0)
            {
                string path = files[0];
                if (Directory.Exists(path))
                {
                    cmbFolderPath.Text = path;
                }
                else if (File.Exists(path))
                {
                    cmbFolderPath.Text = Path.GetDirectoryName(path);
                }
            }
        }
    }
}