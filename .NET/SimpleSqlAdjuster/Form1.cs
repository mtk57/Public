using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

namespace SimpleSqlAdjuster
{
    public partial class MainForm : Form
    {
        private readonly SqlAdjuster _sqlAdjuster = new SqlAdjuster();
        private readonly UserSettingsService _settingsService = new UserSettingsService();
        private UserSettings _settings;

        public MainForm()
        {
            InitializeComponent();
            btnRun.Click += btnRun_Click;
            Load += MainForm_Load;
            FormClosing += MainForm_FormClosing;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

            _settings = _settingsService.Load();

            if (_settings.WindowWidth > 0 && _settings.WindowHeight > 0)
            {
                Size = new Size(_settings.WindowWidth, _settings.WindowHeight);
            }

            if (_settings.WindowX != 0 || _settings.WindowY != 0)
            {
                StartPosition = FormStartPosition.Manual;
                Location = new Point(_settings.WindowX, _settings.WindowY);
            }

            if (!string.IsNullOrEmpty(_settings.LastBeforeSql))
            {
                txtBeforeSQL.Text = _settings.LastBeforeSql;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (_settings == null)
                {
                    _settings = new UserSettings();
                }

                var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                _settings.LastBeforeSql = txtBeforeSQL.Text;

                _settingsService.Save(
                    _settings,
                    new Size(bounds.Width, bounds.Height),
                    new Point(bounds.X, bounds.Y));
            }
            catch (Exception ex)
            {
                LogService.Log(ex, "設定保存時にエラーが発生しました。");
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                var result = _sqlAdjuster.Process(txtBeforeSQL.Text);
                txtAfterSQL.Text = result;
                if (!string.IsNullOrEmpty(result))
                {
                    txtAfterSQL.SelectionStart = txtAfterSQL.TextLength;
                    txtAfterSQL.ScrollToCaret();
                }
            }
            catch (SqlProcessingException ex)
            {
                var message = ex.ToDisplayMessage();
                txtAfterSQL.Text = message;
                LogService.Log(ex, message);
                MessageBox.Show(this, message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                const string message = "予期しないエラーが発生しました。";
                txtAfterSQL.Text = message;
                LogService.Log(ex, message);
                MessageBox.Show(this, message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void txtAfterSQL_TextChanged ( object sender, EventArgs e )
        {

        }
    }
}
