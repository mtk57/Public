using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SimpleGrep
{
    public partial class MultiKeywordsForm : Form
    {
        public MultiKeywordsForm () : this( string.Empty )
        {
        }

        public MultiKeywordsForm ( string initialText )
        {
            InitializeComponent();
            AcceptButton = btnApply;
            CancelButton = btnCancel;
            btnApply.Click += btnApply_Click;
            btnCancel.Click += btnCancel_Click;
            MultiKeywordsText = initialText;
        }

        public string MultiKeywordsText
        {
            get => txtMultiKeywords.Text;
            set => txtMultiKeywords.Text = value ?? string.Empty;
        }

        public IReadOnlyList<string> Keywords => ParseKeywords();

        private IReadOnlyList<string> ParseKeywords ()
        {
            var separators = new[] { "\r\n", "\n", "\r" };
            var rawText = MultiKeywordsText ?? string.Empty;
            return rawText
                .Split( separators, StringSplitOptions.None )
                .Where( line => !string.IsNullOrWhiteSpace( line ) )
                .ToList();
        }

        private void btnApply_Click ( object sender, EventArgs e )
        {
            if ( Keywords.Count == 0 )
            {
                MessageBox.Show( "検索キーワードを1行以上入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Information );
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click ( object sender, EventArgs e )
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
