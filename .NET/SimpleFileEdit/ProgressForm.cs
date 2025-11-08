using System;
using System.Windows.Forms;

namespace SimpleFileEdit
{
    public partial class ProgressForm : Form
    {
        private Action _cancelAction;

        public ProgressForm ()
        {
            InitializeComponent();
        }

        public void Initialize ( string title, int totalFiles, Action cancelAction )
        {
            Text = title;
            _cancelAction = cancelAction;
            btnStop.Enabled = true;
            progressBar.Minimum = 0;
            progressBar.Maximum = Math.Max(1, totalFiles);
            progressBar.Value = 0;
            UpdateProgress(0, totalFiles);
        }

        public void UpdateProgress ( int processed, int total )
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke((Action)(() => UpdateProgress(processed, total)));
                return;
            }

            progressBar.Maximum = Math.Max(1, total);
            progressBar.Value = Math.Max(progressBar.Minimum, Math.Min(processed, progressBar.Maximum));
            lblStatus.Text = $"{processed:D3}/{total:D3} ファイル";
        }

        private void BtnStop_Click ( object sender, EventArgs e )
        {
            if (_cancelAction == null)
            {
                return;
            }

            btnStop.Enabled = false;
            _cancelAction();
        }
    }
}
