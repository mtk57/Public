using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace SimpleFileEdit
{
    public partial class MainForm : Form
    {
        private const string TargetJava = "Java";
        private const string JavaExtension = ".java";
        private const int FolderHistoryLimit = 10;

        private readonly string _settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SimpleFileEdit.json");
        private AppSettings _settings = new AppSettings();

        public MainForm()
        {
            InitializeComponent();
            btnRefDir.Click += BtnRefDir_Click;
            btnDeleteComment.Click += BtnDeleteComment_Click;
            btnDeleteEmptyRow.Click += BtnDeleteEmptyRow_Click;
            cmbFolderPath.DragEnter += CmbFolderPath_DragEnter;
            cmbFolderPath.DragDrop += CmbFolderPath_DragDrop;
            this.FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            Text = $"{Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            if (cmbTarget.Items.Count > 0 && cmbTarget.SelectedIndex < 0)
            {
                cmbTarget.SelectedIndex = 0;
            }

            chkSearchSubDir.Checked = true;
            LoadSettings();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void BtnRefDir_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "対象フォルダを選択してください";
                if (Directory.Exists(cmbFolderPath.Text))
                {
                    dialog.SelectedPath = cmbFolderPath.Text;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    RememberFolderPath(dialog.SelectedPath);
                }
            }
        }

        private void BtnDeleteComment_Click(object sender, EventArgs e)
        {
            if (!IsJavaSelected())
            {
                MessageBox.Show(this, "コメント削除はJavaのみ対応しています。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProcessFiles(RemoveJavaComments, "コメント削除");
        }

        private void BtnDeleteEmptyRow_Click(object sender, EventArgs e)
        {
            ProcessFiles(RemoveEmptyLines, "空行削除");
        }

        private void CmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) ?? false)
            {
                var paths = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (paths != null && paths.Length > 0 && Directory.Exists(paths[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }

            e.Effect = DragDropEffects.None;
        }

        private void CmbFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            var paths = e.Data?.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || paths.Length == 0)
            {
                return;
            }

            var folderPath = paths[0];
            if (Directory.Exists(folderPath))
            {
                RememberFolderPath(folderPath);
            }
        }

        private void ProcessFiles(Func<string, string> transformer, string operationName)
        {
            if (!TryGetSelectedTargetExtension(out var extension))
            {
                return;
            }

            if (!TryGetValidFolder(out var folderPath))
            {
                return;
            }

            List<string> files;
            try
            {
                var searchOption = chkSearchSubDir.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                files = Directory.EnumerateFiles(folderPath, $"*{extension}", searchOption).ToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"ファイル探索に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (files.Count == 0)
            {
                MessageBox.Show(this, "対象ファイルが見つかりませんでした。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var previousCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            int updatedCount = 0;

            try
            {
                foreach (var file in files)
                {
                    string original;
                    Encoding encoding;
                    try
                    {
                        original = ReadFilePreserveEncoding(file, out encoding);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, $"ファイルの読み込みに失敗しました: {file}\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    var transformed = transformer(original) ?? string.Empty;
                    if (!string.Equals(original, transformed, StringComparison.Ordinal))
                    {
                        try
                        {
                            WriteFilePreserveEncoding(file, transformed, encoding);
                            updatedCount++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(this, $"ファイルの書き込みに失敗しました: {file}\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            finally
            {
                Cursor.Current = previousCursor;
            }

            MessageBox.Show(this, $"{operationName}が完了しました。\n処理対象: {files.Count} ファイル\n更新: {updatedCount} ファイル", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool TryGetValidFolder(out string folderPath)
        {
            folderPath = (cmbFolderPath.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(folderPath))
            {
                MessageBox.Show(this, "フォルダパスを入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                folderPath = Path.GetFullPath(folderPath);
            }
            catch (Exception)
            {
                MessageBox.Show(this, "フォルダパスの形式が正しくありません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Path.IsPathRooted(folderPath))
            {
                MessageBox.Show(this, "フォルダパスは絶対パスで指定してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(folderPath))
            {
                MessageBox.Show(this, "指定されたフォルダが存在しません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            RememberFolderPath(folderPath);
            return true;
        }

        private bool TryGetSelectedTargetExtension(out string extension)
        {
            var target = cmbTarget.SelectedItem as string;
            if (string.Equals(target, TargetJava, StringComparison.OrdinalIgnoreCase))
            {
                extension = JavaExtension;
                return true;
            }

            extension = null;
            MessageBox.Show(this, "未対応の対象が選択されています。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }

        private bool IsJavaSelected()
        {
            var target = cmbTarget.SelectedItem as string;
            return string.Equals(target, TargetJava, StringComparison.OrdinalIgnoreCase);
        }

        private void RememberFolderPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            path = path.Trim();

            if (_settings.FolderHistory == null)
            {
                _settings.FolderHistory = new List<string>();
            }

            _settings.FolderHistory = _settings.FolderHistory
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            _settings.FolderHistory.RemoveAll(p => string.Equals(p, path, StringComparison.OrdinalIgnoreCase));
            _settings.FolderHistory.Insert(0, path);

            if (_settings.FolderHistory.Count > FolderHistoryLimit)
            {
                _settings.FolderHistory = _settings.FolderHistory.Take(FolderHistoryLimit).ToList();
            }

            ApplyHistoryToCombo(_settings.FolderHistory, false);
            cmbFolderPath.Text = path;
        }

        private void ApplyHistoryToCombo(IList<string> history, bool applySelection)
        {
            cmbFolderPath.BeginUpdate();
            cmbFolderPath.Items.Clear();

            if (history != null)
            {
                foreach (var item in history.Where(p => !string.IsNullOrWhiteSpace(p)))
                {
                    cmbFolderPath.Items.Add(item);
                }
            }

            cmbFolderPath.EndUpdate();

            if (applySelection && cmbFolderPath.Items.Count > 0)
            {
                cmbFolderPath.SelectedIndex = 0;
            }
        }

        private void LoadSettings()
        {
            if (!File.Exists(_settingsPath))
            {
                ApplyHistoryToCombo(_settings.FolderHistory, true);
                return;
            }

            try
            {
                var json = File.ReadAllText(_settingsPath, Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                var loaded = serializer.Deserialize<AppSettings>(json);
                if (loaded != null)
                {
                    if (loaded.FolderHistory == null)
                    {
                        loaded.FolderHistory = new List<string>();
                    }
                    _settings = loaded;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"設定ファイルの読み込みに失敗しました: {ex.Message}", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            ApplyHistoryToCombo(_settings.FolderHistory, true);

            if (!string.IsNullOrEmpty(_settings.SelectedTarget))
            {
                var index = cmbTarget.Items.IndexOf(_settings.SelectedTarget);
                if (index >= 0)
                {
                    cmbTarget.SelectedIndex = index;
                }
            }

            chkSearchSubDir.Checked = _settings.SearchSubDirectories;
        }

        private void SaveSettings()
        {
            try
            {
                var currentPath = (cmbFolderPath.Text ?? string.Empty).Trim();
                if (Directory.Exists(currentPath))
                {
                    RememberFolderPath(currentPath);
                }
            }
            catch (Exception)
            {
                // 履歴更新での例外は保存処理を継続する
            }

            _settings.SearchSubDirectories = chkSearchSubDir.Checked;
            _settings.SelectedTarget = cmbTarget.SelectedItem as string ?? TargetJava;

            try
            {
                var serializer = new JavaScriptSerializer();
                var json = serializer.Serialize(_settings);
                File.WriteAllText(_settingsPath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"設定ファイルの保存に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string RemoveJavaComments(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var builder = new StringBuilder(content.Length);
            bool inString = false;
            bool inChar = false;
            bool escapeString = false;
            bool escapeChar = false;
            bool inLineComment = false;
            bool inBlockComment = false;

            for (int i = 0; i < content.Length; i++)
            {
                char current = content[i];
                char next = i + 1 < content.Length ? content[i + 1] : '\0';

                if (inLineComment)
                {
                    if (current == '\r' || current == '\n')
                    {
                        inLineComment = false;
                        builder.Append(current);
                        if (current == '\r' && next == '\n')
                        {
                            builder.Append('\n');
                            i++;
                        }
                    }
                    continue;
                }

                if (inBlockComment)
                {
                    if (current == '*' && next == '/')
                    {
                        inBlockComment = false;
                        i++;
                    }
                    continue;
                }

                if (inString)
                {
                    builder.Append(current);
                    if (escapeString)
                    {
                        escapeString = false;
                    }
                    else if (current == '\\')
                    {
                        escapeString = true;
                    }
                    else if (current == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (inChar)
                {
                    builder.Append(current);
                    if (escapeChar)
                    {
                        escapeChar = false;
                    }
                    else if (current == '\\')
                    {
                        escapeChar = true;
                    }
                    else if (current == '\'')
                    {
                        inChar = false;
                    }

                    continue;
                }

                if (current == '/' && next == '/')
                {
                    inLineComment = true;
                    i++;
                    continue;
                }

                if (current == '/' && next == '*')
                {
                    inBlockComment = true;
                    i++;
                    continue;
                }

                if (current == '"')
                {
                    inString = true;
                    builder.Append(current);
                    continue;
                }

                if (current == '\'')
                {
                    inChar = true;
                    builder.Append(current);
                    continue;
                }

                builder.Append(current);
            }

            return builder.ToString();
        }

        private string RemoveEmptyLines(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var builder = new StringBuilder(content.Length);
            int index = 0;

            while (index < content.Length)
            {
                int lineStart = index;
                int newlineStart = -1;

                while (index < content.Length)
                {
                    char c = content[index];
                    if (c == '\r')
                    {
                        newlineStart = index;
                        index++;
                        if (index < content.Length && content[index] == '\n')
                        {
                            index++;
                        }
                        break;
                    }

                    if (c == '\n')
                    {
                        newlineStart = index;
                        index++;
                        break;
                    }

                    index++;
                }

                int lineEnd = newlineStart >= 0 ? newlineStart : index;
                string line = content.Substring(lineStart, lineEnd - lineStart);
                bool hasNewline = newlineStart >= 0;
                string newline = hasNewline ? content.Substring(lineEnd, index - lineEnd) : string.Empty;

                if (!string.IsNullOrWhiteSpace(line))
                {
                    builder.Append(line);
                    if (hasNewline)
                    {
                        builder.Append(newline);
                    }
                }
            }

            return builder.ToString();
        }

        private class AppSettings
        {
            public List<string> FolderHistory { get; set; } = new List<string>();
            public string SelectedTarget { get; set; } = TargetJava;
            public bool SearchSubDirectories { get; set; } = true;
        }

        private static string ReadFilePreserveEncoding(string path, out Encoding encoding)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var reader = new StreamReader(stream, Encoding.Default, true))
            {
                var content = reader.ReadToEnd();
                encoding = reader.CurrentEncoding;
                return content;
            }
        }

        private static void WriteFilePreserveEncoding(string path, string content, Encoding encoding)
        {
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, encoding))
            {
                writer.Write(content);
            }
        }
    }
}
