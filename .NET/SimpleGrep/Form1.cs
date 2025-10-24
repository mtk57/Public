using System;
using System.Collections.Concurrent; // ConcurrentBagのために追加
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading; // Interlockedのために追加
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class MainForm : Form
    {
        private const string SettingsFileName = "SimpleGrep.settings.json";
        private const int MaxHistoryCount = 10;
        private CancellationTokenSource searchCancellationTokenSource;

        public MainForm()
        {
            InitializeComponent();
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            this.button1.Click += new System.EventHandler(this.btnGrep_Click);
            this.btnExportSakura.Click += new System.EventHandler(this.btnExportSakura_Click);
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            this.btnFileCopy.Click += new System.EventHandler(this.btnFileCopy_Click);
            this.dataGridViewResults.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewResults_CellDoubleClick);
            this.FormClosing += new FormClosingEventHandler(MainForm_FormClosing);
            this.chkMethod.CheckedChanged += new System.EventHandler(this.chkMethod_CheckedChanged);
            this.cmbKeyword.KeyDown += new KeyEventHandler(this.cmbKeyword_KeyDown);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
            dataGridViewResults.Columns.Add("Encoding", "Encoding");
            dataGridViewResults.Columns["Encoding"].Visible = false;
            LoadSettings();
            UpdateMethodColumnVisibility();
            btnCancel.Enabled = false;
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
                    chkMethod.Checked = settings.DeriveMethod;
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
                TagJump = chkTagJump.Checked,
                DeriveMethod = chkMethod.Checked
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
            btnCancel.Enabled = true;
            this.Cursor = Cursors.WaitCursor;

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            labelTime.Text = "";

            bool wasCancelled = false;
            bool searchStarted = false;
            int totalFiles = 0;
            CancellationTokenSource localCancellationTokenSource = null;

            try
            {
                bool searchSubdirectories = chkSearchSubDir.Checked;
                bool caseSensitive = chkCase.Checked;
                bool useRegex = chkUseRegex.Checked;
                bool deriveMethod = chkMethod.Checked;

                var searchOption = searchSubdirectories ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                string[] filesToSearch;
                try
                {
                    filesToSearch = Directory.GetFiles(folderPath, filePattern, searchOption);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"ファイル一覧の取得に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                totalFiles = filesToSearch.Length;

                if (totalFiles == 0)
                {
                    MessageBox.Show("対象ファイルが見つかりません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    progressBar.Value = 0;
                    lblPer.Text = "0 %";
                    return;
                }

                progressBar.Maximum = totalFiles;
                progressBar.Value = 0;
                lblPer.Text = "0 %";

                searchCancellationTokenSource?.Cancel();
                searchCancellationTokenSource?.Dispose();

                localCancellationTokenSource = new CancellationTokenSource();
                searchCancellationTokenSource = localCancellationTokenSource;
                var cancellationToken = localCancellationTokenSource.Token;

                var progress = new Progress<int>(processedCount =>
                {
                    int clamped = Math.Min(processedCount, totalFiles);
                    progressBar.Value = clamped;
                    int percentage = totalFiles == 0 ? 0 : (int)((double)clamped / totalFiles * 100);
                    lblPer.Text = $"{percentage} %";
                });

                IEnumerable<SearchResult> searchResults = null;
                searchStarted = true;

                searchResults = await Task.Run(() => SearchFiles(filesToSearch, grepPattern, progress, caseSensitive, useRegex, deriveMethod, cancellationToken), cancellationToken);
                if (localCancellationTokenSource?.IsCancellationRequested == true)
                {
                    wasCancelled = true;
                }

                if (!wasCancelled && searchResults != null && searchResults.Any())
                {
                    dataGridViewResults.SuspendLayout();
                    var rows = searchResults.Select(r =>
                    {
                        object[] cells =
                        {
                            r.FilePath,
                            r.LineNumber,
                            r.LineText,
                            r.MethodSignature,
                            r.EncodingName
                        };

                        var row = new DataGridViewRow();
                        row.CreateCells(dataGridViewResults, cells);
                        return row;
                    }).ToArray();
                    dataGridViewResults.Rows.AddRange(rows);
                    dataGridViewResults.ResumeLayout();
                }
            }
            catch (OperationCanceledException)
            {
                wasCancelled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                stopwatch.Stop();

                if (searchStarted)
                {
                    labelTime.Text = wasCancelled ? $"Time: {stopwatch.Elapsed:mm\\:ss} (中止)" : $"Time: {stopwatch.Elapsed:mm\\:ss}";

                    if (!wasCancelled)
                    {
                        progressBar.Value = totalFiles;
                        lblPer.Text = "100 %";
                    }
                    else
                    {
                        int percentage = totalFiles == 0 ? 0 : (int)((double)progressBar.Value / totalFiles * 100);
                        lblPer.Text = $"{percentage} % (中止)";
                    }
                }

                button1.Enabled = true;
                btnCancel.Enabled = false;
                this.Cursor = Cursors.Default;

                localCancellationTokenSource?.Dispose();
                searchCancellationTokenSource = null;
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

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (searchCancellationTokenSource != null && !searchCancellationTokenSource.IsCancellationRequested)
            {
                searchCancellationTokenSource.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private void cmbKeyword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                e.Handled = true;
                if (button1.Enabled)
                {
                    button1.PerformClick();
                }
            }
        }

        private void btnFileCopy_Click(object sender, EventArgs e)
        {
            var selectedIndices = dataGridViewResults.SelectedCells.Cast<DataGridViewCell>()
                .Select(cell => cell.RowIndex)
                .Concat(dataGridViewResults.SelectedRows.Cast<DataGridViewRow>().Select(row => row.Index))
                .Distinct()
                .ToList();

            if (selectedIndices.Count == 0)
            {
                MessageBox.Show("ファイルをコピーする行を選択してください。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var files = new StringCollection();
            foreach (var index in selectedIndices)
            {
                if (index < 0 || index >= dataGridViewResults.Rows.Count)
                {
                    continue;
                }

                var cellValue = dataGridViewResults.Rows[index].Cells[0].Value;
                if (cellValue == null)
                {
                    continue;
                }

                var filePath = cellValue.ToString();
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    files.Add(filePath);
                }
            }

            if (files.Count == 0)
            {
                MessageBox.Show("実在するファイルが選択されていません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Clipboard.SetFileDropList(files);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"クリップボードへのコピーに失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private IEnumerable<SearchResult> SearchFiles(string[] filePaths, string grepPattern, IProgress<int> progress, bool caseSensitive, bool useRegex, bool deriveMethod, CancellationToken cancellationToken)
        {
            var results = new ConcurrentBag<SearchResult>(); // スレッドセーフなコレクションを使用
            var regexOptions = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            int processedFileCount = 0;

            try
            {
                // Parallel.ForEach を使ってファイルを並列処理
                Parallel.ForEach(filePaths, new ParallelOptions { CancellationToken = cancellationToken }, filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        Encoding encoding = DetectEncoding(filePath); // ファイルごとにエンコーディングを判定
                        string encodingName = GetEncodingName(encoding);

                        string[] lines;
                        using (var reader = new StreamReader(filePath, encoding))
                        {
                            var content = reader.ReadToEnd();
                            cancellationToken.ThrowIfCancellationRequested();
                            lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                        }

                        IMethodSignatureResolver methodResolver = null;
                        if (deriveMethod)
                        {
                            string extension = Path.GetExtension(filePath);
                            if (string.Equals(extension, ".java", StringComparison.OrdinalIgnoreCase))
                            {
                                methodResolver = new JavaMethodSignatureResolver(lines);
                            }
                            else if (string.Equals(extension, ".cs", StringComparison.OrdinalIgnoreCase))
                            {
                                methodResolver = new CSharpMethodSignatureResolver(lines);
                            }
                        }

                        for (int index = 0; index < lines.Length; index++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            string line = lines[index];
                            int lineNumber = index + 1;
                            bool isMatch = false;

                            if (useRegex)
                            {
                                try
                                {
                                    isMatch = Regex.IsMatch(line, grepPattern, regexOptions);
                                }
                                catch (ArgumentException)
                                {
                                    // 無効な正規表現パターンはスキップ
                                }
                            }
                            else
                            {
                                var comparisonType = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                                isMatch = line.IndexOf(grepPattern, comparisonType) >= 0;
                            }

                            if (isMatch)
                            {
                                string methodSignature = methodResolver?.GetMethodSignature(lineNumber) ?? string.Empty;
                                results.Add(new SearchResult
                                {
                                    FilePath = filePath,
                                    LineNumber = lineNumber,
                                    LineText = line,
                                    MethodSignature = methodSignature,
                                    EncodingName = encodingName
                                });
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Skip file read errors
                    }
                    finally
                    {
                        // 処理済みファイル数をスレッドセーフにインクリメント
                        int currentCount = Interlocked.Increment(ref processedFileCount);

                        // 100ファイルごと、または最後のファイル処理時に進捗を通知
                        if (currentCount % 100 == 0 || currentCount == filePaths.Length)
                        {
                            progress?.Report(Math.Min(currentCount, filePaths.Length));
                        }
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // キャンセル時は例外を表面化させずに処理終了
            }
            catch (AggregateException ex) when (ex.InnerExceptions.All(e => e is OperationCanceledException))
            {
                // Parallel.ForEach が AggregateException でキャンセルを通知するケースに備える
            }
            catch (Exception ex)
            {
                // UIスレッドで例外を表示
                this.Invoke((Action)(() =>
                {
                    MessageBox.Show($"ディレクトリの検索中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }));
            }
            return results;
        }

        private void chkMethod_CheckedChanged(object sender, EventArgs e)
        {
            UpdateMethodColumnVisibility();
        }

        private void UpdateMethodColumnVisibility()
        {
            if (clmMethodSignature != null)
            {
                clmMethodSignature.Visible = chkMethod.Checked;
            }
        }

        // ★★ここから修正★★
        private void btnExportSakura_Click(object sender, EventArgs e)
        {
            if (dataGridViewResults.Rows.Count == 0)
            {
                MessageBox.Show("エクスポートするデータがありません。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string fileName = Path.Combine(AppContext.BaseDirectory, DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".grep");
            try
            {
                using (var writer = new StreamWriter(fileName, false, Encoding.UTF8))
                {
                    foreach (DataGridViewRow row in dataGridViewResults.Rows)
                    {
                        string filePath = row.Cells[0].Value.ToString();
                        string lineNumber = row.Cells[1].Value.ToString();
                        string lineContent = row.Cells[2].Value.ToString();
                        var encodingCell = row.Cells["Encoding"];
                        string encoding = encodingCell != null && encodingCell.Value != null ? encodingCell.Value.ToString() : "UTF-8";
                        
                        writer.WriteLine($"{filePath}({lineNumber},1)  [{encoding}]: {lineContent}");
                    }
                }
                
                string sakuraPath = FindSakuraPath();
                if (sakuraPath != null)
                {
                    try
                    {
                        Process.Start(sakuraPath, $"\"{fileName}\"");
                        MessageBox.Show($"{fileName} に結果を保存し、サクラエディタで開きました。", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルのエクスポートには成功しましたが、サクラエディタの起動に失敗しました。\nエラー: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show($"{fileName} に結果を保存しました。\n(サクラエディタが見つからなかったため、ファイルは自動で開かれませんでした)", "エクスポート完了", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"エクスポート中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // ★★ここまで修正★★

        private string GetEncodingName(Encoding encoding)
        {
            // 日本語(Shift-JIS)のコードページは932
            if (encoding.CodePage == 932)
            {
                return "Shift_JIS";
            }
            // UTF-8 BOM付き と BOMなし を区別せず "UTF-8" として表示
            if (encoding is UTF8Encoding)
            {
                return "UTF-8";
            }
            return encoding.WebName.ToUpper();
        }

        /// <summary>
        /// ファイルのエンコーディングを判別します。
        /// </summary>
        /// <param name="filePath">判別するファイルのパス。</param>
        /// <returns>検出されたエンコーディング。</returns>
        private Encoding DetectEncoding(string filePath)
        {
            byte[] bom = new byte[4];
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.Read(bom, 0, 4);
                }
            }
            catch {
                return Encoding.Default; // ファイルがロックされている場合などはデフォルト
            }

            // BOMで判定
            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
                return new UTF8Encoding(true); // UTF-8 BOMあり
            if (bom[0] == 0xff && bom[1] == 0xfe)
                return Encoding.Unicode; // UTF-16 LE
            if (bom[0] == 0xfe && bom[1] == 0xff)
                return Encoding.BigEndianUnicode; // UTF-16 BE

            // BOMがない場合、ファイルの内容からUTF-8かShift_JISかを推定
            byte[] fileBytes;
            try
            {
                 fileBytes = File.ReadAllBytes(filePath);
            }
            catch {
                return Encoding.Default;
            }

            if (IsUtf8(fileBytes))
            {
                return new UTF8Encoding(false); // UTF-8 BOMなし
            }

            // UTF-8でなければShift_JISとみなす
            return Encoding.GetEncoding("Shift_JIS");
        }

        /// <summary>
        /// バイト配列が有効なUTF-8シーケンスであるかを確認します。
        /// </summary>
        private bool IsUtf8(byte[] bytes)
        {
            int i = 0;
            while (i < bytes.Length)
            {
                if (bytes[i] < 0x80) // 1バイト文字 (ASCII)
                {
                    i++;
                    continue;
                }

                if (bytes[i] < 0xC2) return false; // 不正な開始バイト

                int extraBytes = 0;
                if (bytes[i] < 0xE0) extraBytes = 1; // 2バイト文字
                else if (bytes[i] < 0xF0) extraBytes = 2; // 3バイト文字
                else if (bytes[i] < 0xF5) extraBytes = 3; // 4バイト文字
                else return false;

                if (i + extraBytes >= bytes.Length) return false;

                // 後続バイトのチェック
                for (int j = 1; j <= extraBytes; j++)
                {
                    if (bytes[i + j] < 0x80 || bytes[i + j] > 0xBF) return false;
                }
                i += (extraBytes + 1);
            }
            return true;
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

        private sealed class SearchResult
        {
            public string FilePath { get; set; }
            public int LineNumber { get; set; }
            public string LineText { get; set; }
            public string MethodSignature { get; set; }
            public string EncodingName { get; set; }
        }

        private interface IMethodSignatureResolver
        {
            string GetMethodSignature(int lineNumber);
        }

        private sealed class JavaMethodSignatureResolver : IMethodSignatureResolver
        {
            private static readonly HashSet<string> MethodModifiers = new HashSet<string>(StringComparer.Ordinal)
            {
                "public","protected","private","static","final","abstract","synchronized","native","strictfp","default"
            };

            private static readonly HashSet<string> DisallowedMethodNames = new HashSet<string>(StringComparer.Ordinal)
            {
                "if","for","while","switch","catch","try","return","new","else","do"
            };

            private readonly List<MethodSpan> _methodSpans;

            public JavaMethodSignatureResolver(string[] lines)
            {
                _methodSpans = BuildSpans(lines);
            }

            public string GetMethodSignature(int lineNumber)
            {
                if (_methodSpans.Count == 0)
                {
                    return string.Empty;
                }

                int low = 0;
                int high = _methodSpans.Count - 1;
                while (low <= high)
                {
                    int mid = (low + high) / 2;
                    var span = _methodSpans[mid];
                    if (lineNumber < span.StartLine)
                    {
                        high = mid - 1;
                    }
                    else if (lineNumber > span.EndLine)
                    {
                        low = mid + 1;
                    }
                    else
                    {
                        return span.Signature;
                    }
                }

                return string.Empty;
            }

            private static List<MethodSpan> BuildSpans(string[] lines)
            {
                var spans = new List<MethodSpan>();
                var methodStack = new Stack<MethodSpanBuilder>();
                bool inBlockComment = false;
                bool capturingSignature = false;
                int signatureStartLine = -1;
                var signatureBuilder = new StringBuilder();
                int braceDepth = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    string rawLine = lines[i] ?? string.Empty;
                    string withoutComments = RemoveComments(rawLine, ref inBlockComment);
                    string braceSafeLine = RemoveStringLiterals(withoutComments);
                    string trimmed = withoutComments.Trim();

                    if (!capturingSignature && StartsWithAnnotation(trimmed))
                    {
                        // アノテーションは次行の署名判定に影響させない
                    }
                    else
                    {
                        if (!capturingSignature && ContainsMethodToken(trimmed))
                        {
                            capturingSignature = true;
                            signatureBuilder.Clear();
                            signatureBuilder.Append(trimmed);
                            signatureStartLine = i + 1;
                        }
                        else if (capturingSignature && !string.IsNullOrWhiteSpace(trimmed))
                        {
                            signatureBuilder.Append(" ").Append(trimmed);
                        }

                        if (capturingSignature)
                        {
                            string candidate = signatureBuilder.ToString();
                            if (candidate.Contains(";") && !candidate.Contains("{"))
                            {
                                // インターフェース等の宣言は対象外
                                capturingSignature = false;
                            }
                            else if (candidate.Contains("{"))
                            {
                                if (TryParseJavaMethodSignature(candidate, out string signature))
                                {
                                    methodStack.Push(new MethodSpanBuilder
                                    {
                                        Signature = signature,
                                        StartLine = signatureStartLine,
                                        StartDepth = braceDepth
                                    });
                                }
                                capturingSignature = false;
                            }
                        }
                    }

                    braceDepth += CountChar(braceSafeLine, '{');
                    braceDepth -= CountChar(braceSafeLine, '}');
                    if (braceDepth < 0)
                    {
                        braceDepth = 0;
                    }

                    while (methodStack.Count > 0 && braceDepth <= methodStack.Peek().StartDepth)
                    {
                        var builder = methodStack.Pop();
                        builder.EndLine = i + 1;
                        spans.Add(builder.ToSpan());
                    }
                }

                while (methodStack.Count > 0)
                {
                    var builder = methodStack.Pop();
                    builder.EndLine = lines.Length;
                    spans.Add(builder.ToSpan());
                }

                spans.Sort((a, b) => a.StartLine.CompareTo(b.StartLine));
                return spans;
            }

            private static bool ContainsMethodToken(string trimmedLine)
            {
                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    return false;
                }

                if (trimmedLine.StartsWith("@", StringComparison.Ordinal))
                {
                    return false;
                }

                int parenIndex = trimmedLine.IndexOf('(');
                if (parenIndex < 0)
                {
                    return false;
                }

                string before = trimmedLine.Substring(0, parenIndex);
                if (before.IndexOf(" class ", StringComparison.Ordinal) >= 0 ||
                    before.IndexOf(" interface ", StringComparison.Ordinal) >= 0 ||
                    before.IndexOf(" enum ", StringComparison.Ordinal) >= 0)
                {
                    return false;
                }

                if (before.EndsWith("class", StringComparison.Ordinal) ||
                    before.EndsWith("interface", StringComparison.Ordinal) ||
                    before.EndsWith("enum", StringComparison.Ordinal))
                {
                    return false;
                }

                return true;
            }

            private static int CountChar(string text, char target)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return 0;
                }

                int count = 0;
                foreach (char c in text)
                {
                    if (c == target)
                    {
                        count++;
                    }
                }
                return count;
            }

            private static string RemoveStringLiterals(string text)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return string.Empty;
                }

                var sb = new StringBuilder(text.Length);
                bool inString = false;
                char delimiter = '\0';

                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];

                    if (!inString && (c == '"' || c == '\''))
                    {
                        inString = true;
                        delimiter = c;
                        sb.Append(' ');
                        continue;
                    }

                    if (inString)
                    {
                        if (c == delimiter && (i == 0 || text[i - 1] != '\\'))
                        {
                            inString = false;
                            delimiter = '\0';
                        }
                        sb.Append(' ');
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }

                return sb.ToString();
            }

            private static string RemoveComments(string line, ref bool inBlockComment)
            {
                if (string.IsNullOrEmpty(line))
                {
                    return string.Empty;
                }

                var sb = new StringBuilder();
                bool inString = false;
                char stringDelimiter = '\0';

                for (int i = 0; i < line.Length; i++)
                {
                    char current = line[i];

                    if (inBlockComment)
                    {
                        if (current == '*' && i + 1 < line.Length && line[i + 1] == '/')
                        {
                            inBlockComment = false;
                            i++;
                        }
                        continue;
                    }

                    if (inString)
                    {
                        sb.Append(current);
                        if (current == stringDelimiter && (i == 0 || line[i - 1] != '\\'))
                        {
                            inString = false;
                            stringDelimiter = '\0';
                        }
                        continue;
                    }

                    if (current == '"' || current == '\'')
                    {
                        inString = true;
                        stringDelimiter = current;
                        sb.Append(current);
                        continue;
                    }

                    if (current == '/' && i + 1 < line.Length)
                    {
                        char next = line[i + 1];
                        if (next == '/')
                        {
                            break;
                        }
                        if (next == '*')
                        {
                            inBlockComment = true;
                            i++;
                            continue;
                        }
                    }

                    sb.Append(current);
                }

                return sb.ToString();
            }

            private static bool TryParseJavaMethodSignature(string candidate, out string signature)
            {
                signature = string.Empty;
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return false;
                }

                int braceIndex = candidate.IndexOf('{');
                if (braceIndex >= 0)
                {
                    candidate = candidate.Substring(0, braceIndex);
                }

                candidate = candidate.Trim();

                int parenOpen = candidate.IndexOf('(');
                int parenClose = candidate.LastIndexOf(')');
                if (parenOpen < 0 || parenClose <= parenOpen)
                {
                    return false;
                }

                string before = candidate.Substring(0, parenOpen).Trim();
                if (string.IsNullOrEmpty(before) || before.Contains("="))
                {
                    return false;
                }

                string parameters = candidate.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();

                var tokens = SplitTokens(before);
                tokens.RemoveAll(t => t.StartsWith("@", StringComparison.Ordinal));

                if (tokens.Count == 0)
                {
                    return false;
                }

                var coreTokens = tokens.Where(t => !MethodModifiers.Contains(t)).ToList();
                if (coreTokens.Count == 0)
                {
                    return false;
                }

                string methodName = coreTokens[coreTokens.Count - 1];
                if (!IsValidIdentifier(methodName) || DisallowedMethodNames.Contains(methodName))
                {
                    return false;
                }

                string returnType = coreTokens.Count > 1 ? string.Join(" ", coreTokens.Take(coreTokens.Count - 1)) : string.Empty;

                string normalizedReturnType = NormalizeWhitespace(returnType);
                string normalizedParameters = NormalizeWhitespace(parameters);
                string normalizedModifiers = NormalizeWhitespace(string.Join(" ", tokens.Where(MethodModifiers.Contains)));

                var signatureParts = new List<string>();
                if (!string.IsNullOrEmpty(normalizedModifiers))
                {
                    signatureParts.Add(normalizedModifiers);
                }
                if (!string.IsNullOrEmpty(normalizedReturnType))
                {
                    signatureParts.Add(normalizedReturnType);
                }
                signatureParts.Add(methodName);

                signature = $"{string.Join(" ", signatureParts)}({normalizedParameters})".Trim();
                return true;
            }

            private static List<string> SplitTokens(string input)
            {
                return input.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }

            private static bool IsValidIdentifier(string token)
            {
                if (string.IsNullOrEmpty(token))
                {
                    return false;
                }

                if (!(char.IsLetter(token[0]) || token[0] == '_' || token[0] == '$'))
                {
                    return false;
                }

                for (int i = 1; i < token.Length; i++)
                {
                    char c = token[i];
                    if (!(char.IsLetterOrDigit(c) || c == '_' || c == '$'))
                    {
                        return false;
                    }
                }

                return true;
            }

            private static string NormalizeWhitespace(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return text?.Trim() ?? string.Empty;
                }

                var sb = new StringBuilder();
                bool previousWhitespace = false;
                foreach (char c in text)
                {
                    if (char.IsWhiteSpace(c))
                    {
                        if (!previousWhitespace)
                        {
                            sb.Append(' ');
                            previousWhitespace = true;
                        }
                    }
                    else
                    {
                        sb.Append(c);
                        previousWhitespace = false;
                    }
                }

                return sb.ToString().Trim();
            }

            private sealed class MethodSpan
            {
                public int StartLine { get; set; }
                public int EndLine { get; set; }
                public string Signature { get; set; }
            }

            private sealed class MethodSpanBuilder
            {
                public string Signature { get; set; }
                public int StartLine { get; set; }
                public int StartDepth { get; set; }
                public int EndLine { get; set; }

                public MethodSpan ToSpan()
                {
                    return new MethodSpan
                    {
                        StartLine = StartLine,
                        EndLine = EndLine,
                        Signature = Signature
                    };
                }
            }

            private static bool StartsWithAnnotation(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                return text[0] == '@';
            }
        }

        private sealed class CSharpMethodSignatureResolver : IMethodSignatureResolver
        {
            private static readonly HashSet<string> MethodModifiers = new HashSet<string>(StringComparer.Ordinal)
            {
                "public","protected","private","internal","static","virtual","override","abstract","sealed","async","unsafe","extern","new","partial"
            };

            private static readonly HashSet<string> ControlKeywords = new HashSet<string>(StringComparer.Ordinal)
            {
                "if","for","foreach","while","switch","catch","using","lock","return","else","do","checked","unchecked","fixed","try"
            };

            private static readonly HashSet<string> DisallowedMethodNames = new HashSet<string>(StringComparer.Ordinal)
            {
                "if","for","foreach","while","switch","catch","using","lock","return","else","do","checked","unchecked","fixed","try","case","default","nameof"
            };

            private readonly List<MethodSpan> _methodSpans;

            public CSharpMethodSignatureResolver(string[] lines)
            {
                _methodSpans = BuildSpans(lines);
            }

            public string GetMethodSignature(int lineNumber)
            {
                if (_methodSpans.Count == 0)
                {
                    return string.Empty;
                }

                int low = 0;
                int high = _methodSpans.Count - 1;
                while (low <= high)
                {
                    int mid = (low + high) / 2;
                    var span = _methodSpans[mid];
                    if (lineNumber < span.StartLine)
                    {
                        high = mid - 1;
                    }
                    else if (lineNumber > span.EndLine)
                    {
                        low = mid + 1;
                    }
                    else
                    {
                        return span.Signature;
                    }
                }

                return string.Empty;
            }

            private static List<MethodSpan> BuildSpans(string[] lines)
            {
                var spans = new List<MethodSpan>();
                var methodStack = new Stack<MethodSpanBuilder>();
                bool inBlockComment = false;
                bool capturingSignature = false;
                int signatureStartLine = -1;
                var signatureBuilder = new StringBuilder();
                int braceDepth = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    string rawLine = lines[i] ?? string.Empty;
                    string withoutComments = RemoveComments(rawLine, ref inBlockComment);
                    string braceSafeLine = RemoveStringLiterals(withoutComments);
                    string trimmed = withoutComments.Trim();

                    if (!capturingSignature && ContainsMethodToken(trimmed))
                    {
                        capturingSignature = true;
                        signatureBuilder.Clear();
                        signatureBuilder.Append(trimmed);
                        signatureStartLine = i + 1;
                    }
                    else if (capturingSignature && !string.IsNullOrWhiteSpace(trimmed))
                    {
                        signatureBuilder.Append(" ").Append(trimmed);
                    }

                    if (capturingSignature)
                    {
                        string candidate = signatureBuilder.ToString();
                        bool hasBrace = candidate.IndexOf('{') >= 0;
                        bool hasArrow = ContainsLambdaArrow(candidate);

                        if (candidate.IndexOf(';') >= 0 && !hasBrace && !hasArrow)
                        {
                            capturingSignature = false;
                        }
                        else if (hasBrace || hasArrow)
                        {
                            if (TryParseCSharpMethodSignature(candidate, out string signature, out bool hasBlockBody, out bool expressionBody))
                            {
                                if (hasBlockBody)
                                {
                                    methodStack.Push(new MethodSpanBuilder
                                    {
                                        Signature = signature,
                                        StartLine = signatureStartLine,
                                        StartDepth = braceDepth
                                    });
                                }
                                else if (expressionBody)
                                {
                                    spans.Add(new MethodSpan
                                    {
                                        Signature = signature,
                                        StartLine = signatureStartLine,
                                        EndLine = i + 1
                                    });
                                }
                            }
                            capturingSignature = false;
                        }
                    }

                    braceDepth += CountChar(braceSafeLine, '{');
                    braceDepth -= CountChar(braceSafeLine, '}');
                    if (braceDepth < 0)
                    {
                        braceDepth = 0;
                    }

                    while (methodStack.Count > 0 && braceDepth <= methodStack.Peek().StartDepth)
                    {
                        var builder = methodStack.Pop();
                        builder.EndLine = i + 1;
                        spans.Add(builder.ToSpan());
                    }
                }

                while (methodStack.Count > 0)
                {
                    var builder = methodStack.Pop();
                    builder.EndLine = lines.Length;
                    spans.Add(builder.ToSpan());
                }

                spans.Sort((a, b) => a.StartLine.CompareTo(b.StartLine));
                return spans;
            }

            private static bool ContainsMethodToken(string trimmedLine)
            {
                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    return false;
                }

                string normalized = StripLeadingAttributes(trimmedLine);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    return false;
                }

                if (normalized.StartsWith("#", StringComparison.Ordinal))
                {
                    return false;
                }

                int parenIndex = normalized.IndexOf('(');
                if (parenIndex < 0)
                {
                    return false;
                }

                string before = normalized.Substring(0, parenIndex);
                if (string.IsNullOrWhiteSpace(before))
                {
                    return false;
                }

                string lowerBefore = before.ToLowerInvariant();
                if (lowerBefore.Contains(" class ") ||
                    lowerBefore.Contains(" struct ") ||
                    lowerBefore.Contains(" interface ") ||
                    lowerBefore.Contains(" record ") ||
                    lowerBefore.Contains(" enum ") ||
                    lowerBefore.Contains(" delegate "))
                {
                    return false;
                }

                string trimmedBefore = before.TrimStart();
                foreach (var keyword in ControlKeywords)
                {
                    if (trimmedBefore.StartsWith(keyword + " ", StringComparison.Ordinal) ||
                        trimmedBefore.StartsWith(keyword + "\t", StringComparison.Ordinal) ||
                        trimmedBefore.StartsWith(keyword + "(", StringComparison.Ordinal) ||
                        trimmedBefore.Equals(keyword, StringComparison.Ordinal))
                    {
                        return false;
                    }
                }

                return true;
            }

            private static bool TryParseCSharpMethodSignature(string candidate, out string signature, out bool hasBlockBody, out bool expressionBody)
            {
                signature = string.Empty;
                hasBlockBody = false;
                expressionBody = false;

                if (string.IsNullOrWhiteSpace(candidate))
                {
                    return false;
                }

                int braceIndex = candidate.IndexOf('{');
                int arrowIndex = IndexOfLambdaArrow(candidate);
                if (arrowIndex >= 0 && (braceIndex < 0 || arrowIndex < braceIndex))
                {
                    expressionBody = true;
                    hasBlockBody = false;
                }
                else if (braceIndex >= 0)
                {
                    hasBlockBody = true;
                }

                string working = candidate;
                if (expressionBody && arrowIndex >= 0)
                {
                    working = working.Substring(0, arrowIndex);
                }
                else if (hasBlockBody && braceIndex >= 0)
                {
                    working = working.Substring(0, braceIndex);
                }

                working = working.Trim();

                int parenOpen = working.IndexOf('(');
                if (parenOpen < 0)
                {
                    return false;
                }

                int parenClose = FindMatchingParenthesis(working, parenOpen);
                if (parenClose < 0)
                {
                    return false;
                }

                string before = working.Substring(0, parenOpen).Trim();
                string parameters = working.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();
                string after = working.Substring(parenClose + 1).Trim();

                string whereClause = string.Empty;
                if (!string.IsNullOrEmpty(after))
                {
                    int whereIndex = IndexOfWhereClause(after);
                    if (whereIndex >= 0)
                    {
                        whereClause = after.Substring(whereIndex).Trim();
                    }
                }

                string strippedBefore = StripLeadingAttributes(before).Trim();
                if (string.IsNullOrEmpty(strippedBefore))
                {
                    return false;
                }

                var tokens = SplitTokens(strippedBefore);
                if (tokens.Count == 0)
                {
                    return false;
                }

                var coreTokens = tokens.Where(t => !MethodModifiers.Contains(t)).ToList();
                if (coreTokens.Count == 0)
                {
                    return false;
                }

                int operatorIndex = coreTokens.IndexOf("operator");
                bool hasConversionKeyword = operatorIndex > 0 && (coreTokens[operatorIndex - 1] == "implicit" || coreTokens[operatorIndex - 1] == "explicit");
                int nameStartIndex = operatorIndex >= 0
                    ? (hasConversionKeyword ? operatorIndex - 1 : operatorIndex)
                    : coreTokens.Count - 1;

                if (nameStartIndex < 0 || nameStartIndex >= coreTokens.Count)
                {
                    return false;
                }

                var methodNameParts = coreTokens.Skip(nameStartIndex).ToList();
                if (methodNameParts.Count == 0)
                {
                    return false;
                }

                string methodName = string.Join(" ", methodNameParts);
                if (operatorIndex < 0)
                {
                    if (DisallowedMethodNames.Contains(methodName) || !IsValidMethodName(methodName))
                    {
                        return false;
                    }
                }

                var returnTokens = coreTokens.Take(nameStartIndex).ToList();
                string returnType = string.Join(" ", returnTokens);

                string normalizedReturnType = NormalizeWhitespace(returnType);
                string normalizedParameters = NormalizeWhitespace(parameters);

                var signatureParts = new List<string>();
                if (!string.IsNullOrEmpty(normalizedReturnType))
                {
                    signatureParts.Add(normalizedReturnType);
                }
                signatureParts.Add(methodName);

                string signatureText = $"{string.Join(" ", signatureParts)}({normalizedParameters})".Trim();
                if (!string.IsNullOrEmpty(whereClause))
                {
                    signatureText = $"{signatureText} {NormalizeWhitespace(whereClause)}";
                }

                signature = signatureText;
                return true;
            }

            private static string StripLeadingAttributes(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return string.Empty;
                }

                string working = text.TrimStart();
                while (working.StartsWith("[", StringComparison.Ordinal))
                {
                    int end = FindMatchingBracket(working, 0);
                    if (end < 0)
                    {
                        break;
                    }

                    working = working.Substring(end + 1).TrimStart();
                }
                return working;
            }

            private static int FindMatchingBracket(string text, int startIndex)
            {
                int depth = 0;
                for (int i = startIndex; i < text.Length; i++)
                {
                    char c = text[i];
                    if (c == '[')
                    {
                        depth++;
                    }
                    else if (c == ']')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            return i;
                        }
                    }
                }
                return -1;
            }

            private static int FindMatchingParenthesis(string text, int openIndex)
            {
                int depth = 0;
                for (int i = openIndex; i < text.Length; i++)
                {
                    char c = text[i];
                    if (c == '(')
                    {
                        depth++;
                    }
                    else if (c == ')')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            return i;
                        }
                    }
                }
                return -1;
            }

            private static int IndexOfWhereClause(string text)
            {
                int index = text.IndexOf("where", StringComparison.Ordinal);
                while (index >= 0)
                {
                    bool beforeOk = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                    int afterIndex = index + "where".Length;
                    bool afterOk = afterIndex >= text.Length || char.IsWhiteSpace(text[afterIndex]);
                    if (beforeOk && afterOk)
                    {
                        return index;
                    }
                    index = text.IndexOf("where", index + 1, StringComparison.Ordinal);
                }
                return -1;
            }

            private static bool ContainsLambdaArrow(string text)
            {
                return IndexOfLambdaArrow(text) >= 0;
            }

            private static int IndexOfLambdaArrow(string text)
            {
                return text.IndexOf("=>", StringComparison.Ordinal);
            }

            private static int CountChar(string text, char target)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return 0;
                }

                int count = 0;
                foreach (char c in text)
                {
                    if (c == target)
                    {
                        count++;
                    }
                }
                return count;
            }

            private static string RemoveStringLiterals(string text)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return string.Empty;
                }

                var sb = new StringBuilder(text.Length);
                bool inString = false;
                char delimiter = '\0';
                bool verbatim = false;

                for (int i = 0; i < text.Length; i++)
                {
                    char current = text[i];
                    if (!inString)
                    {
                        if (current == '"' || current == '\'')
                        {
                            inString = true;
                            delimiter = current;
                            verbatim = false;
                            if (current == '"')
                            {
                                bool prevAt = i > 0 && text[i - 1] == '@';
                                bool prevDollar = i > 0 && text[i - 1] == '$';
                                bool prevAtBefore = i > 1 && text[i - 2] == '@' && prevDollar;
                                bool prevDollarBefore = i > 1 && text[i - 2] == '$' && prevAt;
                                if (prevAt || prevAtBefore)
                                {
                                    verbatim = true;
                                }
                            }
                            sb.Append(' ');
                        }
                        else
                        {
                            sb.Append(current);
                        }
                    }
                    else
                    {
                        sb.Append(' ');
                        if (current == delimiter)
                        {
                            if (delimiter == '"' && verbatim)
                            {
                                if (i + 1 < text.Length && text[i + 1] == '"')
                                {
                                    sb.Append(' ');
                                    i++;
                                }
                                else
                                {
                                    inString = false;
                                    verbatim = false;
                                    delimiter = '\0';
                                }
                            }
                            else if (i == 0 || text[i - 1] != '\\')
                            {
                                inString = false;
                                delimiter = '\0';
                            }
                        }
                    }
                }

                return sb.ToString();
            }

            private static string RemoveComments(string line, ref bool inBlockComment)
            {
                if (string.IsNullOrEmpty(line))
                {
                    return string.Empty;
                }

                var sb = new StringBuilder();
                bool inString = false;
                char stringDelimiter = '\0';
                bool verbatimString = false;

                for (int i = 0; i < line.Length; i++)
                {
                    char current = line[i];

                    if (inBlockComment)
                    {
                        if (current == '*' && i + 1 < line.Length && line[i + 1] == '/')
                        {
                            inBlockComment = false;
                            i++;
                        }
                        continue;
                    }

                    if (inString)
                    {
                        sb.Append(current);
                        if (current == stringDelimiter)
                        {
                            if (stringDelimiter == '"' && verbatimString)
                            {
                                if (i + 1 < line.Length && line[i + 1] == '"')
                                {
                                    sb.Append(line[i + 1]);
                                    i++;
                                }
                                else
                                {
                                    inString = false;
                                    verbatimString = false;
                                    stringDelimiter = '\0';
                                }
                            }
                            else if (i == 0 || line[i - 1] != '\\')
                            {
                                inString = false;
                                stringDelimiter = '\0';
                            }
                        }
                        continue;
                    }

                    if (current == '"' || current == '\'')
                    {
                        inString = true;
                        stringDelimiter = current;
                        verbatimString = false;
                        if (current == '"')
                        {
                            bool prevAt = i > 0 && line[i - 1] == '@';
                            bool prevDollar = i > 0 && line[i - 1] == '$';
                            bool prevAtBefore = i > 1 && line[i - 2] == '@' && prevDollar;
                            bool prevDollarBefore = i > 1 && line[i - 2] == '$' && prevAt;
                            if (prevAt || prevAtBefore)
                            {
                                verbatimString = true;
                            }
                        }
                        sb.Append(current);
                        continue;
                    }

                    if (current == '/' && i + 1 < line.Length)
                    {
                        char next = line[i + 1];
                        if (next == '/')
                        {
                            break;
                        }
                        if (next == '*')
                        {
                            inBlockComment = true;
                            i++;
                            continue;
                        }
                    }

                    sb.Append(current);
                }

                return sb.ToString();
            }

            private static List<string> SplitTokens(string input)
            {
                return input.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }

            private static bool IsValidMethodName(string token)
            {
                if (string.IsNullOrEmpty(token))
                {
                    return false;
                }

                string working = token.Trim();
                if (working.StartsWith("@", StringComparison.Ordinal))
                {
                    working = working.Substring(1);
                }

                if (string.IsNullOrEmpty(working))
                {
                    return false;
                }

                if (working.StartsWith("~", StringComparison.Ordinal))
                {
                    return IsValidIdentifier(working.Substring(1));
                }

                if (working.Contains("."))
                {
                    var parts = working.Split('.');
                    return parts.All(IsValidIdentifierWithGenerics);
                }

                return IsValidIdentifierWithGenerics(working);
            }

            private static bool IsValidIdentifierWithGenerics(string text)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return false;
                }

                int genericIndex = text.IndexOf('<');
                if (genericIndex >= 0)
                {
                    if (!text.EndsWith(">", StringComparison.Ordinal))
                    {
                        return false;
                    }
                    string prefix = text.Substring(0, genericIndex);
                    return IsValidIdentifier(prefix);
                }

                return IsValidIdentifier(text);
            }

            private static bool IsValidIdentifier(string text)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return false;
                }

                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (i == 0)
                    {
                        if (!(char.IsLetter(c) || c == '_'))
                        {
                            return false;
                        }
                    }
                    else
                    {
                        if (!(char.IsLetterOrDigit(c) || c == '_'))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }

            private static string NormalizeWhitespace(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return text?.Trim() ?? string.Empty;
                }

                var sb = new StringBuilder();
                bool previousWhitespace = false;
                foreach (char c in text)
                {
                    if (char.IsWhiteSpace(c))
                    {
                        if (!previousWhitespace)
                        {
                            sb.Append(' ');
                            previousWhitespace = true;
                        }
                    }
                    else
                    {
                        sb.Append(c);
                        previousWhitespace = false;
                    }
                }

                return sb.ToString().Trim();
            }

            private sealed class MethodSpan
            {
                public int StartLine { get; set; }
                public int EndLine { get; set; }
                public string Signature { get; set; }
            }

            private sealed class MethodSpanBuilder
            {
                public string Signature { get; set; }
                public int StartLine { get; set; }
                public int StartDepth { get; set; }
                public int EndLine { get; set; }

                public MethodSpan ToSpan()
                {
                    return new MethodSpan
                    {
                        StartLine = StartLine,
                        EndLine = EndLine,
                        Signature = Signature
                    };
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

            [DataMember]
            public bool DeriveMethod { get; set; }
        }
    }
}
