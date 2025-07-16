using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class MainForm : Form
    {
        private const string SettingsFileName = "SimpleGrep.settings.json";
        private const int MaxHistoryCount = 10;

        public MainForm()
        {
            InitializeComponent();
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            this.button1.Click += new System.EventHandler(this.btnGrep_Click);
            this.btnExportSakura.Click += new System.EventHandler(this.btnExportSakura_Click);
            this.dataGridViewResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewResults_CellDoubleClick);
            this.FormClosing += new FormClosingEventHandler(MainForm_FormClosing);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
            dataGridViewResults.Columns.Add("Encoding", "Encoding");
            dataGridViewResults.Columns["Encoding"].Visible = false;
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void dataGridViewResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            try
            {
                string filePath = dataGridViewResults.Rows[e.RowIndex].Cells[0].Value.ToString();

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("ファイルが見つかりません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (chkTagJump.Checked)
                {
                    string lineNumber = dataGridViewResults.Rows[e.RowIndex].Cells[1].Value.ToString();
                    string sakuraPath = FindSakuraPath();

                    if (sakuraPath != null)
                    {
                        Process.Start(sakuraPath, $"-Y={lineNumber} \"{filePath}\"");
                    }
                    else
                    {
                        Process.Start(filePath);
                    }
                }
                else
                {
                    string directoryPath = Path.GetDirectoryName(filePath);
                    Process.Start("explorer.exe", $"\"{directoryPath}\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作を実行できませんでした: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string FindSakuraPath()
        {
            try
            {
                var configPath = ConfigurationManager.AppSettings["SakuraEditorPath"];
                if (!string.IsNullOrEmpty(configPath) && File.Exists(configPath))
                {
                    return configPath;
                }
            }
            catch (ConfigurationErrorsException)
            {
            }

            string[] searchPaths = {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "sakura", "sakura.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "sakura", "sakura.exe")
            };

            foreach (var path in searchPaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }
            
            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (pathVar != null)
            {
                foreach (var p in pathVar.Split(Path.PathSeparator))
                {
                    var fullPath = Path.Combine(p, "sakura.exe");
                    if (File.Exists(fullPath))
                        return fullPath;
                }
            }

            return null;
        }


        private void LoadSettings()
        {
            if (!File.Exists(SettingsFileName)) return;

            try
            {
                using (var stream = new FileStream(SettingsFileName, FileMode.Open, FileAccess.Read))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    var settings = (AppSettings)serializer.ReadObject(stream);

                    LoadHistory(cmbFolderPath, settings.FolderPathHistory);
                    LoadHistory(comboBox1, settings.FilePatternHistory);
                    LoadHistory(cmbKeyword, settings.GrepPatternHistory);

                    chkSearchSubDir.Checked = settings.SearchSubDir;
                    chkCase.Checked = settings.CaseSensitive;
                    chkUseRegex.Checked = settings.UseRegex;
                    chkTagJump.Checked = settings.TagJump;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSettings()
        {
            var settings = new AppSettings
            {
                FolderPathHistory = GetHistory(cmbFolderPath),
                FilePatternHistory = GetHistory(comboBox1),
                GrepPatternHistory = GetHistory(cmbKeyword),
                SearchSubDir = chkSearchSubDir.Checked,
                CaseSensitive = chkCase.Checked,
                UseRegex = chkUseRegex.Checked,
                TagJump = chkTagJump.Checked
            };

            try
            {
                using (var stream = new FileStream(SettingsFileName, FileMode.Create, FileAccess.Write))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    serializer.WriteObject(stream, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadHistory(ComboBox comboBox, List<string> history)
        {
            if (history != null && history.Any())
            {
                comboBox.Items.AddRange(history.ToArray());
                comboBox.Text = history.First();
            }
        }

        private List<string> GetHistory(ComboBox comboBox)
        {
            var history = new List<string>();
            if (!string.IsNullOrWhiteSpace(comboBox.Text))
            {
                history.Add(comboBox.Text);
            }
            history.AddRange(comboBox.Items.Cast<string>().Where(item => item != comboBox.Text));
            return history.Distinct().Take(MaxHistoryCount).ToList();
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

            UpdateHistory(cmbFolderPath, folderPath);
            UpdateHistory(comboBox1, filePattern);
            UpdateHistory(cmbKeyword, grepPattern);

            dataGridViewResults.Rows.Clear();
            button1.Enabled = false;
            this.Cursor = Cursors.WaitCursor;
    
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            labelTime.Text = ""; // 表示をリセット

            var searchOption = chkSearchSubDir.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            string[] filesToSearch = Directory.GetFiles(folderPath, filePattern, searchOption);
            int totalFiles = filesToSearch.Length;

            if (totalFiles == 0)
            {
                MessageBox.Show("対象ファイルが見つかりません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button1.Enabled = true;
                this.Cursor = Cursors.Default;
                stopwatch.Stop(); // ストップウォッチを止める
                return;
            }

            progressBar.Maximum = totalFiles;
            progressBar.Value = 0;
            lblPer.Text = "0 %";

            try
            {
                await Task.Run(() =>
                {
                    SearchFiles(filesToSearch, grepPattern);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                stopwatch.Stop();
                labelTime.Text = $"Time: {stopwatch.Elapsed.TotalSeconds:F3} sec";
        
                button1.Enabled = true;
                this.Cursor = Cursors.Default;
                progressBar.Value = progressBar.Maximum;
                lblPer.Text = "100 %";
            }
        }
        
        private void UpdateHistory(ComboBox comboBox, string newItem)
        {
            var items = comboBox.Items.Cast<string>().ToList();
            items.Remove(newItem);
            items.Insert(0, newItem);
            comboBox.Items.Clear();
            comboBox.Items.AddRange(items.Take(MaxHistoryCount).ToArray());
            comboBox.Text = newItem;
        }

        private void SearchFiles(string[] filePaths, string grepPattern)
        {
            try
            {
                var regexOptions = chkCase.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
                
                int processedFileCount = 0;
                int totalFiles = filePaths.Length;

                foreach (var filePath in filePaths)
                {
                    try
                    {
                        string encodingName = "UTF-8"; // Default
                        using (var reader = new StreamReader(filePath, true))
                        {
                            encodingName = GetEncodingName(reader.CurrentEncoding);
                            string line;
                            int lineNumber = 1;
                            while ((line = reader.ReadLine()) != null)
                            {
                                bool isMatch = false;
                                if (chkUseRegex.Checked)
                                {
                                    try
                                    {
                                        isMatch = Regex.IsMatch(line, grepPattern, regexOptions);
                                    }
                                    catch (ArgumentException)
                                    {
                                        // Invalid regex pattern, skip
                                    }
                                }
                                else
                                {
                                    var comparisonType = chkCase.Checked ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                                    isMatch = line.IndexOf(grepPattern, comparisonType) >= 0;
                                }

                                if (isMatch)
                                {
                                    this.Invoke((Action)(() =>
                                    {
                                        dataGridViewResults.Rows.Add(filePath, lineNumber, line, encodingName);
                                    }));
                                }
                                lineNumber++;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Skip file read errors
                    }
                    finally
                    {
                        processedFileCount++;
                        int percentage = (int)((double)processedFileCount / totalFiles * 100);

                        // UIスレッドでプログレスバーとラベルを更新
                        this.Invoke((Action)(() =>
                        {
                            progressBar.Value = processedFileCount;
                            lblPer.Text = $"{percentage} %";
                        }));
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

        private void btnExportSakura_Click(object sender, EventArgs e)
        {
            if (dataGridViewResults.Rows.Count == 0)
            {
                MessageBox.Show("エクスポートするデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                string fileName = DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".grep";
                using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
                {
                    foreach (DataGridViewRow row in dataGridViewResults.Rows)
                    {
                        string filePath = row.Cells[0].Value.ToString();
                        string lineNumber = row.Cells[1].Value.ToString();
                        string lineContent = row.Cells[2].Value.ToString();
                        string encoding = row.Cells.Count > 3 && row.Cells[3].Value != null ? row.Cells[3].Value.ToString() : "UTF-8";
                        
                        writer.WriteLine($"{filePath}({lineNumber},1)  [{encoding}]: {lineContent}");
                    }
                }
                MessageBox.Show($"{fileName} に結果を保存しました。", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetEncodingName(Encoding encoding)
        {
            if (encoding.Equals(Encoding.UTF8))
                return "UTF-8";
            if (encoding.Equals(Encoding.Unicode))
                return "UTF-16";
            if (encoding.Equals(Encoding.Default))
                return "Shift_JIS"; // Or appropriate default
            return encoding.WebName.ToUpper();
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

        [DataContract]
        private class AppSettings
        {
            [DataMember]
            public List<string> FolderPathHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> FilePatternHistory { get; set; } = new List<string>();
            [DataMember]
            public List<string> GrepPatternHistory { get; set; } = new List<string>();
            [DataMember]
            public bool SearchSubDir { get; set; }
            [DataMember]
            public bool CaseSensitive { get; set; }
            [DataMember]
            public bool UseRegex { get; set; }
            
            [DataMember]
            public bool TagJump { get; set; } = true;
        }
    }
}