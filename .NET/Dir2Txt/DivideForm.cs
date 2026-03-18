using System;
using System.Windows.Forms;

namespace Dir2Txt
{
    public partial class DivideForm : Form
    {
        private const int MARGIN = 10;

        private readonly string _text;
        private readonly int _chunkSize;
        private readonly int _totalChunks;
        private int _copyCount;

        public DivideForm ( string text, int divideLength )
        {
            InitializeComponent();

            _text = text;
            _chunkSize = divideLength - MARGIN;
            _totalChunks = (int) Math.Ceiling( (double) _text.Length / _chunkSize );
            _copyCount = 0;

            UpdateCountLabel();

            btnCopyToClipboard.Click += BtnCopyToClipboard_Click;
            btnClose.Click += BtnClose_Click;
        }

        private void BtnCopyToClipboard_Click ( object sender, EventArgs e )
        {
            var start = _copyCount * _chunkSize;
            var length = Math.Min( _chunkSize, _text.Length - start );
            var chunk = _text.Substring( start, length );

            Clipboard.SetText( chunk );
            _copyCount++;
            UpdateCountLabel();

            if ( _copyCount >= _totalChunks )
            {
                btnCopyToClipboard.Enabled = false;
            }
        }

        private void BtnClose_Click ( object sender, EventArgs e )
        {
            Close();
        }

        private void UpdateCountLabel ()
        {
            lblCount.Text = $"{_copyCount}/{_totalChunks}";
        }
    }
}
