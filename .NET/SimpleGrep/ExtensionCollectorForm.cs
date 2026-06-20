using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class ExtensionCollectorForm : Form
    {
        private CancellationTokenSource cancellationTokenSource;

        public ExtensionCollectorForm() : this(string.Empty)
        {
        }

        public ExtensionCollectorForm(string initialFolderPath)
        {
            InitializeComponent();
            txtFolderPath.Text = initialFolderPath ?? string.Empty;
            btnStart.Click += btnStart_Click;
            btnCancel.Click += btnCancel_Click;
            FormClosing += ExtensionCollectorForm_FormClosing;
            btnCancel.Enabled = false;
        }

        private async void btnStart_Click(object sender, EventArgs e)
        {
            string folderPath = txtFolderPath.Text;
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                MessageBox.Show("フォルダパスを正しく入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            cancellationTokenSource?.Cancel();
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = new CancellationTokenSource();
            var cancellationToken = cancellationTokenSource.Token;

            btnStart.Enabled = false;
            btnCancel.Enabled = true;
            txtResult.Text = string.Empty;
            lblStatus.Text = "収集中...";
            Cursor = Cursors.WaitCursor;

            try
            {
                var extensions = await Task.Run(() => CollectExtensions(folderPath, cancellationToken), cancellationToken);
                txtResult.Text = string.Join("/", extensions);
                lblStatus.Text = string.Format("{0} 件", extensions.Count);
            }
            catch (OperationCanceledException)
            {
                lblStatus.Text = "中止";
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("拡張子の収集中にエラーが発生しました: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "エラー";
            }
            finally
            {
                btnStart.Enabled = true;
                btnCancel.Enabled = false;
                Cursor = Cursors.Default;

                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (cancellationTokenSource != null && !cancellationTokenSource.IsCancellationRequested)
            {
                cancellationTokenSource.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private void ExtensionCollectorForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            cancellationTokenSource?.Cancel();
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = null;
        }

        private static IReadOnlyList<string> CollectExtensions(string folderPath, CancellationToken cancellationToken)
        {
            var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string filePath in EnumerateFilesSafely(folderPath, cancellationToken))
            {
                cancellationToken.ThrowIfCancellationRequested();

                string extension = Path.GetExtension(filePath);
                if (string.IsNullOrWhiteSpace(extension))
                {
                    continue;
                }

                extensions.Add(extension.TrimStart('.'));
            }

            return extensions
                .OrderBy(extension => extension, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static IEnumerable<string> EnumerateFilesSafely(string folderPath, CancellationToken cancellationToken)
        {
            var directories = new Stack<string>();
            directories.Push(folderPath);

            while (directories.Count > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                string currentDirectory = directories.Pop();

                string[] files;
                try
                {
                    files = Directory.GetFiles(currentDirectory);
                }
                catch (UnauthorizedAccessException)
                {
                    continue;
                }
                catch (IOException)
                {
                    continue;
                }

                foreach (string file in files)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    yield return file;
                }

                string[] subDirectories;
                try
                {
                    subDirectories = Directory.GetDirectories(currentDirectory);
                }
                catch (UnauthorizedAccessException)
                {
                    continue;
                }
                catch (IOException)
                {
                    continue;
                }

                foreach (string subDirectory in subDirectories)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    directories.Push(subDirectory);
                }
            }
        }
    }
}
