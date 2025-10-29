using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SimpleMethodCallListCreator
{
    public partial class MainForm : Form
    {
        private const int MaxHistoryCount = 10;

        private AppSettings _settings;
        private readonly BindingList<MethodCallDetail> _bindingSource = new BindingList<MethodCallDetail>();
        private List<MethodCallDetail> _sourceResults = new List<MethodCallDetail>();
        private List<MethodCallDetail> _allResults = new List<MethodCallDetail>();

        public MainForm()
        {
            InitializeComponent();
            InitializeGrid();
            HookEvents();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LoadSettings();

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
        }

        private void InitializeGrid()
        {
            dataGridViewResults.AutoGenerateColumns = false;
            dataGridViewResults.DataSource = _bindingSource;
            dataGridViewResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewResults.MultiSelect = true;
            dataGridViewResults.RowHeadersVisible = false;
            dataGridViewResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridViewResults.AllowUserToAddRows = false;
            dataGridViewResults.AllowUserToDeleteRows = false;

            clmFilePath.DataPropertyName = nameof(MethodCallDetail.FilePath);
            clmFileName.DataPropertyName = nameof(MethodCallDetail.FileName);
            clmClassName.DataPropertyName = nameof(MethodCallDetail.ClassName);
            clmCallerMethod.DataPropertyName = nameof(MethodCallDetail.CallerMethod);
            clmCalleeClass.DataPropertyName = nameof(MethodCallDetail.CalleeClass);
            clmCalleeMethod.DataPropertyName = nameof(MethodCallDetail.CalleeMethodName);
            clmCalleeMethodParams.DataPropertyName = nameof(MethodCallDetail.CalleeMethodArguments);
            clmRowNumCalleeMethod.DataPropertyName = nameof(MethodCallDetail.LineNumber);
        }

        private void HookEvents()
        {
            btnBrowse.Click += BtnBrowse_Click;
            btnRun.Click += BtnRun_Click;
            btnExport.Click += BtnExport_Click;
            btnImport.Click += BtnImport_Click;
            btnOther.Click += BtnOther_Click;
            cmbFilePath.DragEnter += CmbFilePath_DragEnter;
            cmbFilePath.DragDrop += CmbFilePath_DragDrop;
            cmbCallerMethod.KeyDown += CmbCallerMethod_KeyDown;
            txtFilePathFilter.TextChanged += FilterTextBox_TextChanged;
            txtFileNameFilter.TextChanged += FilterTextBox_TextChanged;
            txtClassNameFilter.TextChanged += FilterTextBox_TextChanged;
            txtCallerMethodNameFilter.TextChanged += FilterTextBox_TextChanged;
            txtCalleeClassNameFilter.TextChanged += FilterTextBox_TextChanged;
            txtCalleeMethodNameFitter.TextChanged += FilterTextBox_TextChanged;
            txtCalleeMethodParamFilter.TextChanged += FilterTextBox_TextChanged;
            txtRowNumFilter.TextChanged += FilterTextBox_TextChanged;
            button1.Click += BtnEditIgnoreConditions_Click;
            dataGridViewResults.CellDoubleClick += DataGridViewResults_CellDoubleClick;
            dataGridViewResults.KeyDown += DataGridViewResults_KeyDown;
            FormClosing += MainForm_FormClosing;
        }

        private void LoadSettings()
        {
            _settings = SettingsManager.Load();
            if (_settings.RecentFilePaths == null)
            {
                _settings.RecentFilePaths = new List<string>();
            }

            if (_settings.RecentCallerMethods == null)
            {
                _settings.RecentCallerMethods = new List<string>();
            }

            if (_settings.IgnoreConditions == null)
            {
                _settings.IgnoreConditions = new List<IgnoreConditionSetting>();
            }

            if (_settings.IgnoreConditions.Count == 0 && _settings.RecentIgnoreKeywords != null &&
                _settings.RecentIgnoreKeywords.Count > 0)
            {
                foreach (var keyword in _settings.RecentIgnoreKeywords)
                {
                    var trimmed = (keyword ?? string.Empty).Trim();
                    if (trimmed.Length == 0)
                    {
                        continue;
                    }

                    _settings.IgnoreConditions.Add(new IgnoreConditionSetting
                    {
                        Keyword = trimmed,
                        Rule = _settings.IgnoreRule,
                        UseRegex = _settings.UseRegex,
                        MatchCase = _settings.MatchCase
                    });
                }

                _settings.RecentIgnoreKeywords = new List<string>();
                _settings.LastIgnoreKeyword = string.Empty;
                _settings.SelectedIgnoreKeywordIndex = -1;
            }

            cmbFilePath.Items.Clear();
            cmbFilePath.Items.AddRange(_settings.RecentFilePaths.ToArray());
            if (_settings.RecentFilePaths.Count > 0)
            {
                cmbFilePath.Text = _settings.RecentFilePaths[0];
            }

            cmbCallerMethod.Items.Clear();
            foreach (var method in _settings.RecentCallerMethods)
            {
                cmbCallerMethod.Items.Add((method ?? string.Empty).Trim());
            }

            var lastCallerMethod = (_settings.LastCallerMethod ?? string.Empty).Trim();
            if (_settings.SelectedCallerMethodIndex >= 0 &&
                _settings.SelectedCallerMethodIndex < cmbCallerMethod.Items.Count)
            {
                cmbCallerMethod.SelectedIndex = _settings.SelectedCallerMethodIndex;
            }
            else if (!string.IsNullOrEmpty(lastCallerMethod))
            {
                var index = cmbCallerMethod.FindStringExact(lastCallerMethod);
                if (index >= 0)
                {
                    cmbCallerMethod.SelectedIndex = index;
                }
                else
                {
                    cmbCallerMethod.SelectedIndex = -1;
                    cmbCallerMethod.Text = lastCallerMethod;
                }
            }
            else
            {
                cmbCallerMethod.SelectedIndex = -1;
                cmbCallerMethod.Text = string.Empty;
            }

        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Java ファイル (*.java)|*.java|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    cmbFilePath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            ExecuteAnalysis();
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (_bindingSource.Count == 0)
            {
                MessageBox.Show(this, "出力対象がありません。", "情報",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
            var fileName = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".tsv";
            var exportPath = Path.Combine(baseDirectory, fileName);

            try
            {
                WriteResultsToTsv(exportPath);
                MessageBox.Show(this, $"結果を出力しました。\n{exportPath}", "結果出力",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                OpenExportFolder(exportPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"結果出力に失敗しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorLogger.LogException(ex);
            }
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "TSV ファイル (*.tsv)|*.tsv|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                dialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    var imported = ReadResultsFromTsv(dialog.FileName);
                    _sourceResults = new List<MethodCallDetail>(imported);
                    var filtered = ApplyIgnoreConditions(_sourceResults);
                    UpdateGrid(filtered);
                    MessageBox.Show(this, "結果を読み込みました。", "結果入力",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"結果入力に失敗しました。\n{ex.Message}", "エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ErrorLogger.LogException(ex);
                }
            }
        }

        private void BtnOther_Click(object sender, EventArgs e)
        {
            if (_settings == null)
            {
                _settings = new AppSettings();
            }

            using (var dialog = new OtherForm(_settings))
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.ShowDialog(this);
            }

            SaveSettings();
        }

        private void BtnEditIgnoreConditions_Click(object sender, EventArgs e)
        {
            using (var dialog = new IgnoreForm())
            {
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.SetConditions(_settings?.IgnoreConditions);
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _settings.IgnoreConditions = dialog.GetConditions();
                    SaveSettings();
                    ReapplyIgnoreConditions();
                }
            }
        }

        private void ReapplyIgnoreConditions()
        {
            if (_sourceResults == null || _sourceResults.Count == 0)
            {
                UpdateGrid(new List<MethodCallDetail>());
                return;
            }

            var filtered = ApplyIgnoreConditions(_sourceResults);
            UpdateGrid(filtered);
        }

        private void ExecuteAnalysis()
        {
            var filePath = (cmbFilePath.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show(this, "対象ファイルパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show(this, "指定されたファイルが見つかりません。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!string.Equals(Path.GetExtension(filePath), ".java", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "Java ファイル（*.java）のみ指定できます。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var callerMethodFilter = (cmbCallerMethod.Text ?? string.Empty).Trim();

            Cursor = Cursors.WaitCursor;
            try
            {
                var analysisResults = JavaMethodCallAnalyzer.Analyze(filePath);
                if (!string.IsNullOrEmpty(callerMethodFilter))
                {
                    var callerFiltered = new List<MethodCallDetail>();
                    foreach (var detail in analysisResults)
                    {
                        if (string.Equals(detail.CallerMethod, callerMethodFilter, StringComparison.Ordinal))
                        {
                            callerFiltered.Add(detail);
                        }
                    }

                    analysisResults = callerFiltered;
                }

                _sourceResults = new List<MethodCallDetail>(analysisResults);
                var filteredResults = ApplyIgnoreConditions(_sourceResults);
                UpdateGrid(filteredResults);
                UpdateHistories(filePath, callerMethodFilter);
                SaveSettings();
            }
            catch (JavaParseException ex)
            {
                ShowJavaParseError(ex);
                var message = new StringBuilder();
                message.Append("Java解析エラー: ");
                message.Append($"行番号={ex.LineNumber}");
                if (!string.IsNullOrEmpty(ex.InvalidContent))
                {
                    message.Append($", 内容={ex.InvalidContent}");
                }

                ErrorLogger.LogError(message.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"処理中にエラーが発生しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorLogger.LogException(ex);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void ShowJavaParseError(JavaParseException ex)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Javaファイルの解析に失敗しました。");
            builder.AppendLine($"行番号: {ex.LineNumber}");
            if (!string.IsNullOrEmpty(ex.InvalidContent))
            {
                builder.AppendLine($"内容: {ex.InvalidContent}");
            }

            MessageBox.Show(this, builder.ToString(), "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private List<MethodCallDetail> ApplyIgnoreConditions(List<MethodCallDetail> source)
        {
            if (source == null || source.Count == 0)
            {
                return new List<MethodCallDetail>();
            }

            var conditions = _settings?.IgnoreConditions;
            if (conditions == null || conditions.Count == 0)
            {
                return new List<MethodCallDetail>(source);
            }

            var compiled = new List<CompiledIgnoreCondition>();
            foreach (var condition in conditions)
            {
                var keyword = (condition.Keyword ?? string.Empty).Trim();
                if (keyword.Length == 0)
                {
                    continue;
                }

                Regex regex = null;
                if (condition.UseRegex)
                {
                    try
                    {
                        var options = condition.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                        regex = new Regex(keyword, options);
                    }
                    catch (ArgumentException ex)
                    {
                        ErrorLogger.LogError($"除外条件の正規表現が不正です (keyword: {keyword}): {ex.Message}");
                        continue;
                    }
                }

                compiled.Add(new CompiledIgnoreCondition
                {
                    Keyword = keyword,
                    Rule = condition.Rule,
                    Comparison = condition.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase,
                    Regex = regex
                });
            }

            if (compiled.Count == 0)
            {
                return new List<MethodCallDetail>(source);
            }

            var filtered = new List<MethodCallDetail>(source.Count);
            foreach (var item in source)
            {
                var methodName = item.CalleeMethodName ?? string.Empty;
                var exclude = false;
                foreach (var condition in compiled)
                {
                    if (condition.Regex != null)
                    {
                        if (condition.Regex.IsMatch(methodName))
                        {
                            exclude = true;
                            break;
                        }
                    }
                    else if (MatchesRule(methodName, condition.Keyword, condition.Rule, condition.Comparison))
                    {
                        exclude = true;
                        break;
                    }
                }

                if (!exclude)
                {
                    filtered.Add(item);
                }
            }

            return filtered;
        }

        private void WriteResultsToTsv(string exportPath)
        {
            if (string.IsNullOrEmpty(exportPath))
            {
                throw new ArgumentException("出力先パスが不正です。", nameof(exportPath));
            }

            var headers = new[]
            {
                "FilePath",
                "FileName",
                "ClassName",
                "CallerMethod",
                "CalleeClass",
                "CalleeMethod",
                "CalleeMethodArguments",
                "LineNumber"
            };

            using (var writer = new StreamWriter(exportPath, false, Encoding.UTF8))
            {
                writer.WriteLine(string.Join("\t", headers));
                foreach (var detail in _bindingSource)
                {
                    var fields = new[]
                    {
                        EscapeForTsv(detail.FilePath),
                        EscapeForTsv(detail.FileName),
                        EscapeForTsv(detail.ClassName),
                        EscapeForTsv(detail.CallerMethod),
                        EscapeForTsv(detail.CalleeClass),
                        EscapeForTsv(detail.CalleeMethodName),
                        EscapeForTsv(detail.CalleeMethodArguments),
                        detail.LineNumber.ToString(CultureInfo.InvariantCulture)
                    };

                    writer.WriteLine(string.Join("\t", fields));
                }
            }
        }

        private List<MethodCallDetail> ReadResultsFromTsv(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                throw new ArgumentException("読み込み対象のパスが不正です。", nameof(path));
            }

            if (!File.Exists(path))
            {
                throw new FileNotFoundException("指定されたファイルが存在しません。", path);
            }

            var lines = ReadAllLinesWithEncoding(path);
            if (lines.Length == 0)
            {
                return new List<MethodCallDetail>();
            }

            var results = new List<MethodCallDetail>();
            var startIndex = 0;
            if (IsHeaderLine(lines[0]))
            {
                startIndex = 1;
            }

            for (var i = startIndex; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var columns = line.Split('\t');
                if (columns.Length < 8)
                {
                    throw new InvalidDataException($"列数が不足しています。(行番号: {i + 1})");
                }

                var filePath = columns[0].Trim();
                var className = columns[2].Trim();
                var callerMethod = columns[3].Trim();
                var calleeClass = columns[4].Trim();
                var calleeMethod = columns[5].Trim();
                var calleeArguments = columns[6].Trim();

                if (!int.TryParse(columns[7].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var lineNumber))
                {
                    throw new InvalidDataException($"行番号の値が不正です。(行番号: {i + 1})");
                }

                if (string.IsNullOrEmpty(filePath))
                {
                    throw new InvalidDataException($"ファイルパスが空です。(行番号: {i + 1})");
                }

                var detail = new MethodCallDetail(filePath, className, callerMethod, calleeClass,
                    calleeMethod, calleeArguments, lineNumber);
                results.Add(detail);
            }

            return results;
        }

        private bool IsHeaderLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            var expectedHeaders = new[]
            {
                "FilePath",
                "FileName",
                "ClassName",
                "CallerMethod",
                "CalleeClass",
                "CalleeMethod",
                "CalleeMethodArguments",
                "LineNumber"
            };

            var columns = line.Split('\t');
            if (columns.Length != expectedHeaders.Length)
            {
                return false;
            }

            for (var i = 0; i < expectedHeaders.Length; i++)
            {
                if (!string.Equals(columns[i].Trim(), expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private string EscapeForTsv(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(value.Length);
            foreach (var ch in value)
            {
                if (ch == '\t')
                {
                    builder.Append(' ');
                }
                else if (ch == '\r' || ch == '\n')
                {
                    builder.Append(' ');
                }
                else
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private string[] ReadAllLinesWithEncoding(string path)
        {
            var encodings = new[]
            {
                new UTF8Encoding(false, true),
                Encoding.GetEncoding("Shift_JIS")
            };

            foreach (var encoding in encodings)
            {
                try
                {
                    return ReadAllLinesInternal(path, encoding);
                }
                catch (DecoderFallbackException)
                {
                    // 次のエンコーディング候補を試す
                }
                catch (ArgumentException)
                {
                    // 次のエンコーディング候補を試す
                }
            }

            return File.ReadAllLines(path);
        }

        private string[] ReadAllLinesInternal(string path, Encoding encoding)
        {
            var lines = new List<string>();
            using (var reader = new StreamReader(path, encoding, detectEncodingFromByteOrderMarks: true))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    lines.Add(line);
                }
            }

            return lines.ToArray();
        }

        private bool MatchesRule(string target, string keyword, IgnoreRule rule, StringComparison comparison)
        {
            if (target == null)
            {
                target = string.Empty;
            }

            switch (rule)
            {
                case IgnoreRule.StartsWith:
                    return target.StartsWith(keyword, comparison);
                case IgnoreRule.EndsWith:
                    return target.EndsWith(keyword, comparison);
                case IgnoreRule.Contains:
                    return target.IndexOf(keyword, comparison) >= 0;
                case IgnoreRule.Exact:
                    return string.Equals(target, keyword, comparison);
                default:
                    return false;
            }
        }

        private sealed class CompiledIgnoreCondition
        {
            public string Keyword { get; set; }
            public IgnoreRule Rule { get; set; }
            public StringComparison Comparison { get; set; }
            public Regex Regex { get; set; }
        }

        private void UpdateGrid(List<MethodCallDetail> results)
        {
            _allResults = results != null ? new List<MethodCallDetail>(results) : new List<MethodCallDetail>();
            ApplyTextFilters();
        }

        private void ApplyTextFilters()
        {
            var filePathFilter = (txtFilePathFilter.Text ?? string.Empty).Trim();
            var fileNameFilter = (txtFileNameFilter.Text ?? string.Empty).Trim();
            var classNameFilter = (txtClassNameFilter.Text ?? string.Empty).Trim();
            var callerMethodFilter = (txtCallerMethodNameFilter.Text ?? string.Empty).Trim();
            var calleeClassFilter = (txtCalleeClassNameFilter.Text ?? string.Empty).Trim();
            var calleeMethodFilter = (txtCalleeMethodNameFitter.Text ?? string.Empty).Trim();
            var calleeParamFilter = (txtCalleeMethodParamFilter.Text ?? string.Empty).Trim();
            var rowNumberFilter = (txtRowNumFilter.Text ?? string.Empty).Trim();

            _bindingSource.Clear();
            if (_allResults == null || _allResults.Count == 0)
            {
                return;
            }

            foreach (var detail in _allResults)
            {
                if (!MatchesFilter(detail.FilePath, filePathFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.FileName, fileNameFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.ClassName, classNameFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.CallerMethod, callerMethodFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.CalleeClass, calleeClassFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.CalleeMethodName, calleeMethodFilter))
                {
                    continue;
                }

                if (!MatchesFilter(detail.CalleeMethodArguments, calleeParamFilter))
                {
                    continue;
                }

                var lineNumberText = detail.LineNumber.ToString(CultureInfo.InvariantCulture);
                if (!MatchesFilter(lineNumberText, rowNumberFilter))
                {
                    continue;
                }

                _bindingSource.Add(detail);
            }
        }

        private bool MatchesFilter(string source, string filter)
        {
            if (string.IsNullOrEmpty(filter))
            {
                return true;
            }

            source = source ?? string.Empty;
            return source.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void UpdateHistories(string filePath, string callerMethod)
        {
            if (_settings == null)
            {
                _settings = new AppSettings();
            }

            UpdateHistoryList(_settings.RecentFilePaths, filePath, true);
            RefreshComboItems(cmbFilePath, _settings.RecentFilePaths, filePath);

            if (!string.IsNullOrEmpty(callerMethod))
            {
                if (_settings.RecentCallerMethods == null)
                {
                    _settings.RecentCallerMethods = new List<string>();
                }

                UpdateHistoryList(_settings.RecentCallerMethods, callerMethod, false);
                RefreshComboItems(cmbCallerMethod, _settings.RecentCallerMethods, callerMethod);
                var index = cmbCallerMethod.FindStringExact(callerMethod);
                if (index >= 0)
                {
                    cmbCallerMethod.SelectedIndex = index;
                    _settings.SelectedCallerMethodIndex = index;
                }
                else
                {
                    cmbCallerMethod.SelectedIndex = -1;
                    cmbCallerMethod.Text = callerMethod;
                    _settings.SelectedCallerMethodIndex = -1;
                }
            }
            else
            {
                _settings.SelectedCallerMethodIndex = -1;
            }
        }

        private void UpdateHistoryList(List<string> history, string value, bool caseInsensitive)
        {
            if (history == null || string.IsNullOrEmpty(value))
            {
                return;
            }

            var comparer = caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            for (var i = history.Count - 1; i >= 0; i--)
            {
                if (comparer.Equals(history[i], value))
                {
                    history.RemoveAt(i);
                }
            }

            history.Insert(0, value);
            if (history.Count > MaxHistoryCount)
            {
                history.RemoveRange(MaxHistoryCount, history.Count - MaxHistoryCount);
            }
        }

        private void RefreshComboItems(ComboBox comboBox, List<string> source, string selectedValue)
        {
            comboBox.BeginUpdate();
            try
            {
                comboBox.Items.Clear();
                comboBox.Items.AddRange(source.ToArray());
                comboBox.Text = selectedValue;
            }
            finally
            {
                comboBox.EndUpdate();
            }
        }

        private void SaveSettings()
        {
            if (_settings == null)
            {
                _settings = new AppSettings();
            }

            _settings.RecentFilePaths = CaptureHistoryFromCombo(cmbFilePath, true);
            var currentCallerMethod = (cmbCallerMethod.Text ?? string.Empty).Trim();
            var callerHistory = CaptureHistoryFromCombo(cmbCallerMethod, false);
            for (var i = callerHistory.Count - 1; i >= 0; i--)
            {
                if (string.Equals(callerHistory[i], currentCallerMethod, StringComparison.Ordinal))
                {
                    callerHistory.RemoveAt(i);
                }
            }

            if (!string.IsNullOrEmpty(currentCallerMethod))
            {
                callerHistory.Insert(0, currentCallerMethod);
            }

            if (callerHistory.Count > MaxHistoryCount)
            {
                callerHistory.RemoveRange(MaxHistoryCount, callerHistory.Count - MaxHistoryCount);
            }

            _settings.RecentCallerMethods = callerHistory;
            _settings.LastCallerMethod = currentCallerMethod;
            RefreshComboItems(cmbCallerMethod, callerHistory, currentCallerMethod);
            var callerIndex = cmbCallerMethod.FindStringExact(currentCallerMethod);
            if (callerIndex >= 0)
            {
                cmbCallerMethod.SelectedIndex = callerIndex;
                _settings.SelectedCallerMethodIndex = callerIndex;
            }
            else
            {
                cmbCallerMethod.SelectedIndex = -1;
                cmbCallerMethod.Text = currentCallerMethod;
                _settings.SelectedCallerMethodIndex = -1;
            }

            try
            {
                SettingsManager.Save(_settings);
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
        }

        private void CmbFilePath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0 && IsJavaFile(files[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
        }

        private void CmbFilePath_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0)
            {
                return;
            }

            var file = files[0];
            if (!IsJavaFile(file))
            {
                MessageBox.Show(this, "Java ファイル（*.java）をドロップしてください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            cmbFilePath.Text = file;
        }

        private bool IsJavaFile(string filePath)
        {
            return string.Equals(Path.GetExtension(filePath), ".java", StringComparison.OrdinalIgnoreCase);
        }

        private void DataGridViewResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _bindingSource.Count)
            {
                return;
            }

            var detail = _bindingSource[e.RowIndex];
            if (detail == null)
            {
                return;
            }

            if (!File.Exists(detail.FilePath))
            {
                MessageBox.Show(this, "対象ファイルが存在しません。", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                LaunchSakura(detail.FilePath, Math.Max(detail.LineNumber, 1));
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"タグジャンプに失敗しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorLogger.LogException(ex);
            }
        }

        private void DataGridViewResults_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                dataGridViewResults.SelectAll();
                e.Handled = true;
                return;
            }

            if (e.Control && e.KeyCode == Keys.C)
            {
                CopySelectionToClipboard();
                e.Handled = true;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void CmbCallerMethod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Delete)
            {
                return;
            }

            var index = cmbCallerMethod.SelectedIndex;
            if (index < 0 || index >= cmbCallerMethod.Items.Count)
            {
                return;
            }

            var removed = cmbCallerMethod.Items[index] as string;
            cmbCallerMethod.Items.RemoveAt(index);
            if (!string.IsNullOrEmpty(removed) && _settings?.RecentCallerMethods != null)
            {
                _settings.RecentCallerMethods.RemoveAll(x => string.Equals(x, removed, StringComparison.Ordinal));
            }

            e.Handled = true;
        }

        private void FilterTextBox_TextChanged(object sender, EventArgs e)
        {
            ApplyTextFilters();
        }

        private void CopySelectionToClipboard()
        {
            if (dataGridViewResults.SelectedRows == null || dataGridViewResults.SelectedRows.Count == 0)
            {
                return;
            }

            var selectedRows = new List<DataGridViewRow>();
            foreach (DataGridViewRow row in dataGridViewResults.SelectedRows)
            {
                if (!row.IsNewRow)
                {
                    selectedRows.Add(row);
                }
            }

            if (selectedRows.Count == 0)
            {
                return;
            }

            selectedRows.Sort((x, y) => x.Index.CompareTo(y.Index));

            var visibleColumns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn column in dataGridViewResults.Columns)
            {
                if (column.Visible)
                {
                    visibleColumns.Add(column);
                }
            }

            if (visibleColumns.Count == 0)
            {
                return;
            }

            var builder = new StringBuilder();
            for (var rowIndex = 0; rowIndex < selectedRows.Count; rowIndex++)
            {
                var row = selectedRows[rowIndex];
                for (var colIndex = 0; colIndex < visibleColumns.Count; colIndex++)
                {
                    var column = visibleColumns[colIndex];
                    var cell = row.Cells[column.Index];
                    var raw = cell?.Value?.ToString() ?? string.Empty;
                    builder.Append(SanitizeForClipboard(raw));
                    if (colIndex < visibleColumns.Count - 1)
                    {
                        builder.Append('\t');
                    }
                }

                if (rowIndex < selectedRows.Count - 1)
                {
                    builder.Append("\r\n");
                }
            }

            try
            {
                Clipboard.SetText(builder.ToString());
            }
            catch (ExternalException ex)
            {
                ErrorLogger.LogException(ex);
            }
        }

        private string SanitizeForClipboard(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(value.Length);
            foreach (var ch in value)
            {
                if (ch == '\r' || ch == '\n' || ch == '\t')
                {
                    builder.Append(' ');
                }
                else
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private void OpenExportFolder(string exportPath)
        {
            try
            {
                var directory = Path.GetDirectoryName(exportPath);
                if (string.IsNullOrEmpty(directory))
                {
                    return;
                }

                if (!Directory.Exists(directory))
                {
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = directory,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
        }

        private void LaunchSakura(string filePath, int lineNumber)
        {
            var sakuraPath = FindSakuraPath();
            if (!string.IsNullOrEmpty(sakuraPath))
            {
                var arguments = $"-Y={lineNumber} \"{filePath}\"";
                Process.Start(new ProcessStartInfo
                {
                    FileName = sakuraPath,
                    Arguments = arguments,
                    UseShellExecute = false
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
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
            catch (ConfigurationErrorsException ex)
            {
                ErrorLogger.LogException(ex);
            }

            var programFilesPaths = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "sakura", "sakura.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "sakura", "sakura.exe")
            };

            foreach (var candidate in programFilesPaths)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            var pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                foreach (var part in pathEnv.Split(Path.PathSeparator))
                {
                    if (string.IsNullOrEmpty(part))
                    {
                        continue;
                    }

                    var candidate = Path.Combine(part, "sakura.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private List<string> CaptureHistoryFromCombo(ComboBox comboBox, bool caseInsensitive, bool allowEmpty = false)
        {
            var result = new List<string>();
            if (comboBox == null)
            {
                return result;
            }

            var comparer = caseInsensitive ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal;
            var seen = new HashSet<string>(comparer);
            foreach (var item in comboBox.Items)
            {
                if (!(item is string text))
                {
                    continue;
                }

                var trimmed = text.Trim();
                if (trimmed.Length == 0)
                {
                    if (!allowEmpty)
                    {
                        continue;
                    }

                    trimmed = string.Empty;
                }

                if (!seen.Add(trimmed))
                {
                    continue;
                }

                result.Add(trimmed);
                if (result.Count >= MaxHistoryCount)
                {
                    break;
                }
            }

            return result;
        }
    }
}
