using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    public partial class MainForm : Form
    {
        private const int MaxHistoryCount = 10;

        private AppSettings _settings;
        private readonly BindingList<MethodCallDetail> _bindingSource = new BindingList<MethodCallDetail>();

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
        }

        private void InitializeGrid()
        {
            dataGridViewResults.AutoGenerateColumns = false;
            dataGridViewResults.DataSource = _bindingSource;
            dataGridViewResults.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewResults.MultiSelect = false;
            dataGridViewResults.RowHeadersVisible = false;
            dataGridViewResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridViewResults.AllowUserToAddRows = false;
            dataGridViewResults.AllowUserToDeleteRows = false;

            clmFilePath.DataPropertyName = nameof(MethodCallDetail.FilePath);
            clmFileName.DataPropertyName = nameof(MethodCallDetail.FileName);
            clmClassName.DataPropertyName = nameof(MethodCallDetail.ClassName);
            clmCallerMethod.DataPropertyName = nameof(MethodCallDetail.CallerMethod);
            clmCalleeClass.DataPropertyName = nameof(MethodCallDetail.CalleeClass);
            clmCalleeMethod.DataPropertyName = nameof(MethodCallDetail.CalleeMethod);
            clmRowNumCalleeMethod.DataPropertyName = nameof(MethodCallDetail.LineNumber);
        }

        private void HookEvents()
        {
            btnBrowse.Click += BtnBrowse_Click;
            btnRun.Click += BtnRun_Click;
            cmbFilePath.DragEnter += CmbFilePath_DragEnter;
            cmbFilePath.DragDrop += CmbFilePath_DragDrop;
            dataGridViewResults.CellDoubleClick += DataGridViewResults_CellDoubleClick;
            FormClosing += MainForm_FormClosing;
        }

        private void LoadSettings()
        {
            _settings = SettingsManager.Load();
            if (_settings.RecentFilePaths == null)
            {
                _settings.RecentFilePaths = new List<string>();
            }

            if (_settings.RecentIgnoreKeywords == null)
            {
                _settings.RecentIgnoreKeywords = new List<string>();
            }

            cmbFilePath.Items.Clear();
            cmbFilePath.Items.AddRange(_settings.RecentFilePaths.ToArray());
            if (_settings.RecentFilePaths.Count > 0)
            {
                cmbFilePath.Text = _settings.RecentFilePaths[0];
            }

            cmbIgnoreKeyword.Items.Clear();
            cmbIgnoreKeyword.Items.AddRange(_settings.RecentIgnoreKeywords.ToArray());
            if (_settings.RecentIgnoreKeywords.Count > 0)
            {
                cmbIgnoreKeyword.Text = _settings.RecentIgnoreKeywords[0];
            }

            chkUseRegex.Checked = _settings.UseRegex;
            chkCase.Checked = _settings.MatchCase;

            var ruleIndex = (int)_settings.IgnoreRule;
            if (ruleIndex < 0 || ruleIndex >= cmbIgnoreRules.Items.Count)
            {
                ruleIndex = 0;
            }

            cmbIgnoreRules.SelectedIndex = ruleIndex;
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

            var ignoreKeyword = (cmbIgnoreKeyword.Text ?? string.Empty).Trim();
            var useRegex = chkUseRegex.Checked;
            var matchCase = chkCase.Checked;
            var ignoreRule = GetSelectedIgnoreRule();

            Regex ignoreRegex = null;
            if (useRegex && !string.IsNullOrEmpty(ignoreKeyword))
            {
                try
                {
                    var options = matchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                    ignoreRegex = new Regex(ignoreKeyword, options);
                }
                catch (ArgumentException ex)
                {
                    MessageBox.Show(this, $"正規表現が不正です。\n{ex.Message}", "入力エラー",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    ErrorLogger.LogError($"除外キーワードの正規表現エラー: {ex.Message}");
                    return;
                }
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                var analysisResults = JavaMethodCallAnalyzer.Analyze(filePath);
                var filteredResults = FilterResults(analysisResults, ignoreKeyword, ignoreRegex, matchCase, ignoreRule, useRegex);
                UpdateGrid(filteredResults);
                UpdateHistories(filePath, ignoreKeyword);
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

        private List<MethodCallDetail> FilterResults(List<MethodCallDetail> source, string ignoreKeyword,
            Regex ignoreRegex, bool matchCase, IgnoreRule rule, bool useRegex)
        {
            if (source == null || source.Count == 0)
            {
                return new List<MethodCallDetail>();
            }

            if (string.IsNullOrEmpty(ignoreKeyword))
            {
                return source;
            }

            var filtered = new List<MethodCallDetail>(source.Count);
            if (useRegex && ignoreRegex != null)
            {
                foreach (var item in source)
                {
                    if (!ignoreRegex.IsMatch(item.CalleeMethod))
                    {
                        filtered.Add(item);
                    }
                }

                return filtered;
            }

            var tokens = ignoreKeyword.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0)
            {
                return source;
            }

            var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (var item in source)
            {
                var exclude = false;
                foreach (var raw in tokens)
                {
                    var keyword = raw.Trim();
                    if (keyword.Length == 0)
                    {
                        continue;
                    }

                    if (MatchesRule(item.CalleeMethod, keyword, rule, comparison))
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

        private bool MatchesRule(string target, string keyword, IgnoreRule rule, StringComparison comparison)
        {
            switch (rule)
            {
                case IgnoreRule.StartsWith:
                    return target.StartsWith(keyword, comparison);
                case IgnoreRule.EndsWith:
                    return target.EndsWith(keyword, comparison);
                case IgnoreRule.Contains:
                    return target.IndexOf(keyword, comparison) >= 0;
                default:
                    return false;
            }
        }

        private void UpdateGrid(List<MethodCallDetail> results)
        {
            _bindingSource.Clear();
            if (results == null)
            {
                return;
            }

            foreach (var item in results)
            {
                _bindingSource.Add(item);
            }
        }

        private void UpdateHistories(string filePath, string ignoreKeyword)
        {
            if (_settings == null)
            {
                _settings = new AppSettings();
            }

            UpdateHistoryList(_settings.RecentFilePaths, filePath, true);
            RefreshComboItems(cmbFilePath, _settings.RecentFilePaths, filePath);

            if (!string.IsNullOrEmpty(ignoreKeyword))
            {
                UpdateHistoryList(_settings.RecentIgnoreKeywords, ignoreKeyword, false);
                RefreshComboItems(cmbIgnoreKeyword, _settings.RecentIgnoreKeywords, ignoreKeyword);
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

        private IgnoreRule GetSelectedIgnoreRule()
        {
            var index = cmbIgnoreRules.SelectedIndex;
            if (index < 0)
            {
                index = 0;
            }

            return (IgnoreRule)index;
        }

        private void SaveSettings()
        {
            if (_settings == null)
            {
                _settings = new AppSettings();
            }

            _settings.UseRegex = chkUseRegex.Checked;
            _settings.MatchCase = chkCase.Checked;
            _settings.IgnoreRule = GetSelectedIgnoreRule();

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

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
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
    }
}
