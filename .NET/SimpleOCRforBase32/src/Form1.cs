using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tesseract;

namespace SimpleOCRforBase32
{
    public partial class Form1 : Form
    {
        public Form1 ()
        {
            InitializeComponent();
            InitializeEvents();
        }

        private void InitializeEvents ()
        {
            txtImgFilePath.AllowDrop = true;
            txtImgFilePath.DragEnter += TxtImgFilePath_DragEnter;
            txtImgFilePath.DragDrop += TxtImgFilePath_DragDrop;
            btnRefImgFilePath.Click += BtnRefImgFilePath_Click;
            btnStart.Click += BtnStart_Click;
        }

        private void TxtImgFilePath_DragEnter ( object sender, DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
                e.Effect = files != null && files.Length > 0 && File.Exists( files[0] )
                    ? DragDropEffects.Copy
                    : DragDropEffects.None;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TxtImgFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
                if ( files != null && files.Length > 0 && File.Exists( files[0] ) )
                {
                    txtImgFilePath.Text = files[0];
                }
            }
        }

        private void BtnRefImgFilePath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new OpenFileDialog() )
            {
                dialog.Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff|All Files|*.*";
                dialog.Title = "画像ファイルを選択してください";
                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtImgFilePath.Text = dialog.FileName;
                }
            }
        }

        private async void BtnStart_Click ( object sender, EventArgs e )
        {
            var filePath = txtImgFilePath.Text;
            if ( string.IsNullOrWhiteSpace( filePath ) || !File.Exists( filePath ) )
            {
                MessageBox.Show( this, "有効な画像ファイルパスを指定してください。", "ファイル未指定", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            btnStart.Enabled = false;
            txtResultOCR.Clear();

            try
            {
                var ocrText = await Task.Run( () => RunOcr( filePath ) );
                txtResultOCR.Text = ocrText;
            }
            catch ( DirectoryNotFoundException ex )
            {
                MessageBox.Show( this, ex.Message, "tessdata が見つかりません", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"OCR処理でエラーが発生しました。\n{ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
            finally
            {
                btnStart.Enabled = true;
            }
        }

        private string RunOcr ( string filePath )
        {
            var tessdataPath = Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "tessdata" );
            if ( !Directory.Exists( tessdataPath ) )
            {
                throw new DirectoryNotFoundException( $"tessdata フォルダーが見つかりません。以下に配置してください。\n{tessdataPath}" );
            }

            using ( var engine = new TesseractEngine( tessdataPath, "eng", EngineMode.Default ) )
            {
                engine.SetVariable( "tessedit_char_whitelist", "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567" );
                using ( var img = Pix.LoadFromFile( filePath ) )
                using ( var page = engine.Process( img ) )
                {
                    var rawText = page.GetText() ?? string.Empty;
                    var normalized = new string( rawText
                        .Where( c => char.IsLetterOrDigit( c ) || char.IsWhiteSpace( c ) )
                        .ToArray() );
                    normalized = normalized.ToUpperInvariant();
                    return normalized.Replace( "\r", string.Empty ).Trim();
                }
            }
        }
    }
}
