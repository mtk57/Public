using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class MainForm: Form
    {
        private const int MaxHistoryItems = 20;
        private const string SettingsFileName = "SimpleFileSearch.json";

        public MainForm()
        {
            InitializeComponent();

            // ドラッグアンドドロップを有効にする
            this.cmbFolderPath.DragEnter += new DragEventHandler(cmbFolderPath_DragEnter);
            this.cmbFolderPath.DragDrop += new DragEventHandler(cmbFolderPath_DragDrop);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // タイトルにバージョン情報を表示
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        #region ドラッグアンドドロップ

        private void cmbFolderPath_DragEnter(object sender, DragEventArgs e)
        {
            // ドラッグされたものがファイルまたはフォルダの場合、カーソルを変更
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
            // ドロップされたファイルまたはフォルダのパスを取得
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            if (files.Length > 0)
            {
                // 最初のアイテムのパスを使用
                string path = files[0];

                // それがフォルダであるか、ファイルであるかを確認
                if (Directory.Exists(path))
                {
                    // フォルダの場合はそのパスを直接使用
                    cmbFolderPath.Text = path;
                }
                else if (File.Exists(path))
                {
                    // ファイルの場合は、そのファイルが含まれるフォルダのパスを使用
                    cmbFolderPath.Text = Path.GetDirectoryName(path);
                }
            }
        }

        #endregion
    }
}