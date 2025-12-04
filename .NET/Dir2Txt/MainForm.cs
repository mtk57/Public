using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Dir2Txt
{
    public partial class MainForm : Form
    {
        public MainForm ()
        {
            InitializeComponent();
            txtDirPath.AllowDrop = true;
            txtDirPath.DragEnter += PathTextBox_DragEnter;
            txtDirPath.DragDrop += TxtDirPath_DragDrop;
            btnRefDirPath.Click += BtnRefDirPath_Click;
            btnRun.Click += BtnRun_Click;
            btnExtract.Click += BtnExtract_Click;
        }

        private void BtnExtract_Click ( object sender, EventArgs e )
        {
            using ( var form = new ExtractForm( txtOutput.Text ) )
            {
                form.ShowDialog( this );
            }
        }

        private void BtnRun_Click ( object sender, EventArgs e )
        {
            var dirPath = txtDirPath.Text;
            if ( string.IsNullOrWhiteSpace( dirPath ) )
            {
                MessageBox.Show( this, "フォルダパスを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( !Directory.Exists( dirPath ) )
            {
                MessageBox.Show( this, "指定されたフォルダが見つかりません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            try
            {
                txtOutput.Text = BuildDirectoryText( dirPath );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"テキスト化に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        private void BtnRefDirPath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new FolderBrowserDialog() )
            {
                dialog.Description = "フォルダを選択してください";
                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private string BuildDirectoryText ( string rootPath )
        {
            var allFiles = Directory.GetFiles( rootPath, "*", SearchOption.AllDirectories )
                                    .OrderBy( x => x )
                                    .ToList();

            var builder = new StringBuilder();
            foreach ( var file in allFiles )
            {
                if ( IsBinaryFile( file ) )
                {
                    continue;
                }

                var read = ReadFileWithEncoding( file );
                builder.Append( "@@" ).Append( file ).Append( "|" ).AppendLine( read.Encoding.WebName );
                var content = read.Content;
                builder.Append( content );
                if ( !content.EndsWith( Environment.NewLine ) )
                {
                    builder.AppendLine();
                }
            }

            return builder.ToString();
        }

        private (string Content, Encoding Encoding) ReadFileWithEncoding ( string path )
        {
            var bytes = File.ReadAllBytes( path );
            var detected = DetectEncoding( bytes );
            var content = detected.GetString( bytes );
            return ( content, detected );
        }

        private Encoding DetectEncoding ( byte[] bytes )
        {
            if ( bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF )
            {
                return new UTF8Encoding( true );
            }
            if ( bytes.Length >= 2 )
            {
                if ( bytes[0] == 0xFF && bytes[1] == 0xFE )
                {
                    return Encoding.Unicode;
                }
                if ( bytes[0] == 0xFE && bytes[1] == 0xFF )
                {
                    return Encoding.BigEndianUnicode;
                }
            }

            var utf8Strict = new UTF8Encoding( false, true );
            try
            {
                utf8Strict.GetString( bytes );
                return utf8Strict;
            }
            catch
            {
            }

            try
            {
                var sjisStrict = Encoding.GetEncoding( 932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback );
                sjisStrict.GetString( bytes );
                return sjisStrict;
            }
            catch
            {
            }

            return Encoding.Default;
        }

        private void PathTextBox_DragEnter ( object sender, DragEventArgs e )
        {
            e.Effect = e.Data.GetDataPresent( DataFormats.FileDrop ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void TxtDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var paths = e.Data.GetData( DataFormats.FileDrop ) as string[];
            var path = paths?.FirstOrDefault();
            if ( string.IsNullOrEmpty( path ) )
            {
                return;
            }

            if ( Directory.Exists( path ) )
            {
                txtDirPath.Text = path;
            }
            else if ( File.Exists( path ) )
            {
                var dir = Path.GetDirectoryName( path );
                if ( !string.IsNullOrEmpty( dir ) )
                {
                    txtDirPath.Text = dir;
                }
            }
        }

        private bool IsBinaryFile ( string path )
        {
            const int sampleSize = 8000;
            var buffer = new byte[sampleSize];
            try
            {
                using ( var stream = File.OpenRead( path ) )
                {
                    var read = stream.Read( buffer, 0, buffer.Length );
                    for ( int i = 0; i < read; i++ )
                    {
                        var b = buffer[i];
                        if ( b == 0 )
                        {
                            return true;
                        }

                        // 制御コードが多い場合はバイナリとみなす
                        if ( b < 0x09 && b != 0x00 )
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
                return true;
            }

            return false;
        }
    }
}
