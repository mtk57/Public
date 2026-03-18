using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Dir2Txt
{
    public partial class DivideForm : Form
    {
        private const int MARGIN = 10;

        private readonly string _text;
        private readonly List<int> _splitPositions;
        private int _copyCount;

        public DivideForm ( string text, int divideLength )
        {
            InitializeComponent();

            _text = text;
            var chunkSize = divideLength - MARGIN;
            _splitPositions = BuildSplitPositions( chunkSize );
            _copyCount = 0;

            UpdateCountLabel();

            btnCopyToClipboard.Click += BtnCopyToClipboard_Click;
            btnClose.Click += BtnClose_Click;
        }

        private List<int> BuildSplitPositions ( int chunkSize )
        {
            var positions = new List<int>();
            positions.Add( 0 );

            var pos = 0;
            while ( pos < _text.Length )
            {
                var end = pos + chunkSize;
                if ( end >= _text.Length )
                {
                    break;
                }

                var newlinePos = _text.LastIndexOf( '\n', end, end - pos );
                if ( newlinePos > pos )
                {
                    end = newlinePos + 1;
                }

                positions.Add( end );
                pos = end;
            }

            return positions;
        }

        private void BtnCopyToClipboard_Click ( object sender, EventArgs e )
        {
            var start = _splitPositions[_copyCount];
            var end = _copyCount + 1 < _splitPositions.Count ? _splitPositions[_copyCount + 1] : _text.Length;
            var chunk = _text.Substring( start, end - start );

            Clipboard.SetText( chunk );
            _copyCount++;
            UpdateCountLabel();

            if ( _copyCount >= TotalChunks )
            {
                btnCopyToClipboard.Enabled = false;
            }
        }

        private int TotalChunks => _splitPositions.Count;

        private void BtnClose_Click ( object sender, EventArgs e )
        {
            Close();
        }

        private void UpdateCountLabel ()
        {
            lblCount.Text = $"{_copyCount}/{TotalChunks}";
        }
    }
}
