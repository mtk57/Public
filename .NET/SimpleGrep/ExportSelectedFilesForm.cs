using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SimpleGrep.Core;

namespace SimpleGrep
{
    internal sealed class ExportSelectedFilesForm : Form
    {
        private readonly IReadOnlyList<SearchResult> searchResults;
        private readonly TextBox txtOutputFolderPath;
        private readonly Button btnExport;
        private readonly Button btnCancel;

        public ExportSelectedFilesForm(IReadOnlyList<SearchResult> searchResults, string initialOutputFolderPath)
        {
            this.searchResults = searchResults ?? new List<SearchResult>();

            Text = "エクスポート";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 125);

            var label = new Label
            {
                AutoSize = true,
                Location = new Point(20, 18),
                Text = "出力フォルダパス"
            };

            txtOutputFolderPath = new TextBox
            {
                AllowDrop = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Location = new Point(22, 38),
                Size = new Size(516, 19),
                TabIndex = 0,
                Text = initialOutputFolderPath ?? string.Empty
            };
            txtOutputFolderPath.DragEnter += txtOutputFolderPath_DragEnter;
            txtOutputFolderPath.DragDrop += txtOutputFolderPath_DragDrop;

            btnExport = new Button
            {
                Location = new Point(382, 78),
                Size = new Size(75, 29),
                TabIndex = 1,
                Text = "エクスポート",
                UseVisualStyleBackColor = true
            };
            btnExport.Click += btnExport_Click;

            btnCancel = new Button
            {
                Location = new Point(463, 78),
                Size = new Size(75, 29),
                TabIndex = 2,
                Text = "キャンセル",
                UseVisualStyleBackColor = true
            };
            btnCancel.Click += btnCancel_Click;

            Controls.Add(label);
            Controls.Add(txtOutputFolderPath);
            Controls.Add(btnExport);
            Controls.Add(btnCancel);

            AcceptButton = btnExport;
            CancelButton = btnCancel;
        }

        public string OutputFolderPath { get; private set; } = string.Empty;

        private void txtOutputFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            if (HasDirectoryDrop(e))
            {
                e.Effect = DragDropEffects.Copy;
                return;
            }

            e.Effect = DragDropEffects.None;
        }

        private void txtOutputFolderPath_DragDrop(object sender, DragEventArgs e)
        {
            var paths = e.Data.GetData(DataFormats.FileDrop, false) as string[];
            var folderPath = paths?.FirstOrDefault(path => Directory.Exists(path));
            if (!string.IsNullOrEmpty(folderPath))
            {
                txtOutputFolderPath.Text = folderPath;
            }
        }

        private static bool HasDirectoryDrop(DragEventArgs e)
        {
            if (e == null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            var paths = e.Data.GetData(DataFormats.FileDrop, false) as string[];
            return paths != null && paths.Any(path => Directory.Exists(path));
        }

        private async void btnExport_Click(object sender, EventArgs e)
        {
            string outputFolderPath = txtOutputFolderPath.Text;
            if (string.IsNullOrWhiteSpace(outputFolderPath))
            {
                MessageBox.Show("出力フォルダパスを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                outputFolderPath = Path.GetFullPath(outputFolderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("出力フォルダパスが不正です: {0}", ex.Message), "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            btnExport.Enabled = false;
            btnCancel.Enabled = false;
            Cursor = Cursors.WaitCursor;

            try
            {
                var summary = await Task.Run(() => ExportFiles(outputFolderPath));
                OutputFolderPath = outputFolderPath;
                MessageBox.Show(
                    string.Format("エクスポートが完了しました。\nコピー: {0} 件\nスキップ: {1} 件", summary.CopiedCount, summary.SkippedCount),
                    "エクスポート完了",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("エクスポート中にエラーが発生しました: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor = Cursors.Default;
                btnExport.Enabled = true;
                btnCancel.Enabled = true;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private ExportSummary ExportFiles(string outputFolderPath)
        {
            var filePaths = searchResults
                .Select(result => result?.FilePath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(path => Path.GetFullPath(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (filePaths.Count == 0)
            {
                return new ExportSummary(0, 0);
            }

            Directory.CreateDirectory(outputFolderPath);
            int copiedCount = 0;
            int skippedCount = 0;

            foreach (string sourceFilePath in filePaths)
            {
                if (!File.Exists(sourceFilePath))
                {
                    skippedCount++;
                    continue;
                }

                string relativeFilePath = GetFullPathBasedExportPath(sourceFilePath);
                string destinationFilePath = Path.Combine(outputFolderPath, relativeFilePath);
                if (string.Equals(sourceFilePath, Path.GetFullPath(destinationFilePath), StringComparison.OrdinalIgnoreCase))
                {
                    skippedCount++;
                    continue;
                }

                string destinationFolderPath = Path.GetDirectoryName(destinationFilePath);
                if (!string.IsNullOrEmpty(destinationFolderPath))
                {
                    Directory.CreateDirectory(destinationFolderPath);
                }

                File.Copy(sourceFilePath, destinationFilePath, true);
                copiedCount++;
            }

            return new ExportSummary(copiedCount, skippedCount);
        }

        private static string GetFullPathBasedExportPath(string filePath)
        {
            string normalizedFilePath = Path.GetFullPath(filePath);
            string rootPath = Path.GetPathRoot(normalizedFilePath) ?? string.Empty;
            string rootFolderName = rootPath
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                .TrimEnd(Path.VolumeSeparatorChar)
                .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            string pathWithoutRoot = normalizedFilePath.Substring(rootPath.Length)
                .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (string.IsNullOrEmpty(rootFolderName))
            {
                return pathWithoutRoot;
            }

            return Path.Combine(rootFolderName, pathWithoutRoot);
        }

        private sealed class ExportSummary
        {
            public ExportSummary(int copiedCount, int skippedCount)
            {
                CopiedCount = copiedCount;
                SkippedCount = skippedCount;
            }

            public int CopiedCount { get; private set; }
            public int SkippedCount { get; private set; }
        }
    }
}
