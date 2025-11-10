using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SimpleExcelGrep.Forms
{
    public partial class MultiKeywordsForm : Form
    {
        public IReadOnlyList<string> ParsedKeywords { get; private set; } = Array.Empty<string>();

        public string KeywordsText
        {
            get => txtMultiKeywords.Text;
            set => txtMultiKeywords.Text = value ?? string.Empty;
        }

        public MultiKeywordsForm ()
        {
            InitializeComponent();
            btnApply.Click += BtnApply_Click;
            btnCancel.Click += BtnCancel_Click;
        }

        private void BtnApply_Click ( object sender, EventArgs e )
        {
            var keywords = ParseKeywords();
            if (!keywords.Any())
            {
                MessageBox.Show("キーワードを1行以上入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtMultiKeywords.Focus();
                return;
            }

            ParsedKeywords = keywords;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnCancel_Click ( object sender, EventArgs e )
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private List<string> ParseKeywords ()
        {
            return txtMultiKeywords
                .Lines
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .ToList();
        }
    }
}
