using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator.Forms
{
    public partial class InsertTagJumpForm : Form
    {
        private readonly AppSettings _settings;
        private readonly TagJumpEmbeddingService _service = new TagJumpEmbeddingService();

        public InsertTagJumpForm(AppSettings settings)
        {
            _settings = settings ?? new AppSettings();
            InitializeComponent();
            HookEvents();
            LoadSettings();
        }

        private void HookEvents()
        {
            btnRun.Click += BtnRun_Click;
            btnRefMethodListPath.Click += BtnRefMethodListPath_Click;
            btnRefStartSrcFilePath.Click += BtnRefStartSrcFilePath_Click;
            txtStartSrcFilePath.TextChanged += TxtStartSrcFilePath_TextChanged;
            FormClosing += InsertTagJumpForm_FormClosing;

            txtMethodListPath.AllowDrop = true;
            txtMethodListPath.DragEnter += FilePathTextBox_DragEnter;
            txtMethodListPath.DragDrop += TxtMethodListPath_DragDrop;

            txtStartSrcFilePath.AllowDrop = true;
            txtStartSrcFilePath.DragEnter += FilePathTextBox_DragEnter;
            txtStartSrcFilePath.DragDrop += TxtStartSrcFilePath_DragDrop;
        }

        private void LoadSettings()
        {
            txtMethodListPath.Text = _settings.LastTagJumpMethodListPath ?? string.Empty;
            txtStartSrcFilePath.Text = _settings.LastTagJumpSourceFilePath ?? string.Empty;
            txtStartMethod.Text = _settings.LastTagJumpMethod ?? string.Empty;
            if (!string.IsNullOrEmpty(_settings.LastTagJumpPrefix))
            {
                txtTagJumpPrefix.Text = _settings.LastTagJumpPrefix;
            }

            UpdateSourceRootDisplay(txtStartSrcFilePath.Text);
        }

        private void SaveSettings()
        {
            _settings.LastTagJumpMethodListPath = (txtMethodListPath.Text ?? string.Empty).Trim();
            _settings.LastTagJumpSourceFilePath = (txtStartSrcFilePath.Text ?? string.Empty).Trim();
            _settings.LastTagJumpMethod = (txtStartMethod.Text ?? string.Empty).Trim();
            _settings.LastTagJumpPrefix = txtTagJumpPrefix.Text ?? string.Empty;

            SettingsManager.Save(_settings);
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            ExecuteEmbedding();
        }

        private void BtnRefMethodListPath_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "TSV ファイル (*.tsv)|*.tsv|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                var currentPath = (txtMethodListPath.Text ?? string.Empty).Trim();
                if (currentPath.Length > 0)
                {
                    var directory = Path.GetDirectoryName(currentPath);
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        dialog.InitialDirectory = directory;
                    }
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtMethodListPath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRefStartSrcFilePath_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Java ファイル (*.java)|*.java|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                var currentPath = (txtStartSrcFilePath.Text ?? string.Empty).Trim();
                if (currentPath.Length > 0)
                {
                    var directory = Path.GetDirectoryName(currentPath);
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                    {
                        dialog.InitialDirectory = directory;
                    }
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    txtStartSrcFilePath.Text = dialog.FileName;
                }
            }
        }

        private void TxtStartSrcFilePath_TextChanged(object sender, EventArgs e)
        {
            UpdateSourceRootDisplay(txtStartSrcFilePath.Text);
        }

        private void UpdateSourceRootDisplay(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                txtSrcRootDirPath.Text = string.Empty;
                return;
            }

            try
            {
                var directory = Path.GetDirectoryName(path);
                txtSrcRootDirPath.Text = directory ?? string.Empty;
            }
            catch
            {
                txtSrcRootDirPath.Text = string.Empty;
            }
        }

        private void FilePathTextBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TxtMethodListPath_DragDrop(object sender, DragEventArgs e)
        {
            var filePath = ResolveDropFilePath(e);
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }

            if (!string.Equals(Path.GetExtension(filePath), ".tsv", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "TSV ファイルをドロップしてください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtMethodListPath.Text = filePath;
        }

        private void TxtStartSrcFilePath_DragDrop(object sender, DragEventArgs e)
        {
            var filePath = ResolveDropFilePath(e);
            if (string.IsNullOrEmpty(filePath))
            {
                return;
            }

            if (!string.Equals(Path.GetExtension(filePath), ".java", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show(this, "Java ファイルをドロップしてください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            txtStartSrcFilePath.Text = filePath;
        }

        private static string ResolveDropFilePath(DragEventArgs e)
        {
            if (e.Data == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return string.Empty;
            }

            if (!(e.Data.GetData(DataFormats.FileDrop) is string[] paths) || paths.Length == 0)
            {
                return string.Empty;
            }

            var candidate = paths[0];
            if (string.IsNullOrEmpty(candidate) || !File.Exists(candidate))
            {
                return string.Empty;
            }

            return candidate;
        }

        private void InsertTagJumpForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveSettings();
        }

        private void ExecuteEmbedding()
        {
            var methodListPath = (txtMethodListPath.Text ?? string.Empty).Trim();
            var sourceFilePath = (txtStartSrcFilePath.Text ?? string.Empty).Trim();
            var startMethod = (txtStartMethod.Text ?? string.Empty).Trim();
            var prefix = txtTagJumpPrefix.Text ?? string.Empty;

            if (methodListPath.Length == 0)
            {
                MessageBox.Show(this, "メソッドリストのパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (sourceFilePath.Length == 0)
            {
                MessageBox.Show(this, "開始ソースファイルのパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (startMethod.Length == 0)
            {
                MessageBox.Show(this, "開始メソッドを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor = Cursors.WaitCursor;
            try
            {
                var result = _service.Execute(methodListPath, sourceFilePath, startMethod, prefix);
                SaveSettings();

                var hasFailures = result.FailureCount > 0;
                var hasUpdates = result.UpdatedCallCount > 0;

                if (hasUpdates || hasFailures)
                {
                    var message = new StringBuilder();
                    message.AppendLine("処理が完了しました。");
                    message.AppendLine($"更新ファイル数: {result.UpdatedFileCount}");
                    message.AppendLine($"更新箇所数: {result.UpdatedCallCount}");
                    if (hasFailures)
                    {
                        message.AppendLine($"メソッド特定失敗: {result.FailureCount}");
                        message.AppendLine("詳細は error.log を参照してください。");
                        LogFailureDetails(methodListPath, sourceFilePath, startMethod, result);
                    }

                    MessageBox.Show(this, message.ToString(), "処理完了",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(this, "タグジャンプ対象の呼び出しは見つかりませんでした。", "処理完了",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (JavaParseException ex)
            {
                var builder = new StringBuilder();
                builder.AppendLine("Javaファイルの解析に失敗しました。");
                builder.AppendLine($"行番号: {ex.LineNumber}");
                if (!string.IsNullOrEmpty(ex.InvalidContent))
                {
                    builder.AppendLine($"内容: {ex.InvalidContent}");
                }

                MessageBox.Show(this, builder.ToString(), "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var logBuilder = new StringBuilder();
                logBuilder.AppendLine("Javaファイルの解析に失敗しました。");
                logBuilder.AppendLine($"メソッドリスト: {methodListPath}");
                logBuilder.AppendLine($"ソースファイル: {sourceFilePath}");
                logBuilder.AppendLine($"開始メソッド: {startMethod}");
                logBuilder.AppendLine($"行番号: {ex.LineNumber}");
                if (!string.IsNullOrEmpty(ex.InvalidContent))
                {
                    logBuilder.AppendLine($"内容: {ex.InvalidContent}");
                }

                logBuilder.AppendLine("例外詳細:");
                logBuilder.AppendLine(ex.ToString());
                ErrorLogger.LogError(logBuilder.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"タグジャンプ情報の挿入に失敗しました。\n{ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var logBuilder = new StringBuilder();
                logBuilder.AppendLine("タグジャンプ情報の挿入に失敗しました。");
                logBuilder.AppendLine($"メソッドリスト: {methodListPath}");
                logBuilder.AppendLine($"ソースファイル: {sourceFilePath}");
                logBuilder.AppendLine($"開始メソッド: {startMethod}");
                logBuilder.AppendLine("行番号: 不明");
                logBuilder.AppendLine("例外詳細:");
                logBuilder.AppendLine(ex.ToString());
                ErrorLogger.LogError(logBuilder.ToString());
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private static void LogFailureDetails(string methodListPath, string sourceFilePath,
            string startMethod, TagJumpEmbeddingResult result)
        {
            if (result == null || result.FailureCount <= 0)
            {
                return;
            }

            var builder = new StringBuilder();
            builder.AppendLine("メソッド特定に失敗した呼び出しがあります。");
            builder.AppendLine($"メソッドリスト: {methodListPath}");
            builder.AppendLine($"ソースファイル: {sourceFilePath}");
            builder.AppendLine($"開始メソッド: {startMethod}");
            builder.AppendLine($"失敗件数: {result.FailureCount}");
            builder.AppendLine("詳細:");

            foreach (var detail in result.FailureDetails)
            {
                if (detail == null)
                {
                    continue;
                }

                builder.Append("  - ファイル: ");
                builder.Append(detail.FilePath);
                builder.Append(", 行番号: ");
                builder.Append(detail.LineNumber);

                if (!string.IsNullOrEmpty(detail.CallerMethodSignature))
                {
                    builder.Append(", 呼出元: ");
                    builder.Append(detail.CallerMethodSignature);
                }

                if (!string.IsNullOrEmpty(detail.CallExpression))
                {
                    builder.Append(", 呼び出し: ");
                    builder.Append(detail.CallExpression);
                }

                if (!string.IsNullOrEmpty(detail.Reason))
                {
                    var reasonText = detail.Reason.Replace("\r\n", "\n");
                    builder.Append(", 理由: ");
                    if (reasonText.IndexOf('\n') >= 0)
                    {
                        builder.AppendLine();
                        builder.Append(reasonText);
                    }
                    else
                    {
                        builder.Append(reasonText);
                    }
                }

                builder.AppendLine();
            }

            ErrorLogger.LogError(builder.ToString());
        }
    }
}
