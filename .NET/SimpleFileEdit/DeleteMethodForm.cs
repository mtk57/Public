using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SimpleFileSearch
{
    public partial class DeleteMethodForm : Form
    {
        public string Keyword { get; private set; }
        public bool IsRegexEnabled { get; private set; }
        public bool IsNotEnabled { get; private set; }

        public DeleteMethodForm()
        {
            InitializeComponent();
            btnDeleteStart.Click += BtnDeleteStart_Click;
            btnCloseForm.Click += BtnCloseForm_Click;
        }

        private void BtnDeleteStart_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtKeywordForMethodSignature.Text))
            {
                return;
            }

            if (chkEnabledRegExForMethodSignature.Checked)
            {
                try
                {
                    Regex.IsMatch(string.Empty, txtKeywordForMethodSignature.Text);
                }
                catch (ArgumentException)
                {
                    MessageBox.Show(this, "正規表現が不正です。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            var confirm = MessageBox.Show(this,
                "メソッド、import削除を実行します。よろしいですか？",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button1);

            if (confirm != DialogResult.Yes)
            {
                return;
            }

            Keyword = txtKeywordForMethodSignature.Text;
            IsRegexEnabled = chkEnabledRegExForMethodSignature.Checked;
            IsNotEnabled = chkNot.Checked;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCloseForm_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
