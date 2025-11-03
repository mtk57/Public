using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
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

            string preprocessedPath = null;

            try
            {
                preprocessedPath = PreprocessImage( filePath );
            }
            catch
            {
                preprocessedPath = null;
            }

            try
            {
                using ( var engine = new TesseractEngine( tessdataPath, "eng", EngineMode.Default ) )
                {
                    engine.SetVariable( "tessedit_char_whitelist", AllowedCharacters );
                    engine.SetVariable( "preserve_interword_spaces", "1" );
                    engine.SetVariable( "load_system_dawg", "0" );
                    engine.SetVariable( "load_freq_dawg", "0" );
                    engine.SetVariable( "wordrec_enable_assoc", "0" );
                    engine.SetVariable( "language_model_penalty_non_dict_word", "0.15" );
                    engine.SetVariable( "language_model_penalty_non_freq_dict_word", "0.25" );
                    engine.SetVariable( "classify_bln_numeric_mode", "1" );
                    engine.DefaultPageSegMode = PageSegMode.SingleColumn;

                    var candidates = new List<Tuple<string, float>>();

                    if ( !string.IsNullOrEmpty( preprocessedPath ) && File.Exists( preprocessedPath ) )
                    {
                        using ( var img = Pix.LoadFromFile( preprocessedPath ) )
                        using ( var page = engine.Process( img, PageSegMode.SingleColumn ) )
                        {
                            var rawText = page.GetText() ?? string.Empty;
                            candidates.Add( Tuple.Create( NormalizeText( rawText ), page.GetMeanConfidence() ) );
                        }
                    }

                    using ( var img = Pix.LoadFromFile( filePath ) )
                    using ( var page = engine.Process( img, PageSegMode.SingleColumn ) )
                    {
                        var rawText = page.GetText() ?? string.Empty;
                        candidates.Add( Tuple.Create( NormalizeText( rawText ), page.GetMeanConfidence() ) );
                    }

                    var best = candidates
                        .OrderByDescending( c => CountValidLines( c.Item1 ) )
                        .ThenByDescending( c => c.Item2 )
                        .FirstOrDefault();

                    return best?.Item1 ?? string.Empty;
                }
            }
            finally
            {
                if ( !string.IsNullOrEmpty( preprocessedPath ) && File.Exists( preprocessedPath ) )
                {
                    File.Delete( preprocessedPath );
                }
            }
        }

        private static readonly string AllowedCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567-";
        private static readonly HashSet<char> AllowedCharacterSet = new HashSet<char>( AllowedCharacters );
        private static readonly Dictionary<char, char> SoftEquivalents = new Dictionary<char, char>
        {
            { '0', 'O' },
            { '1', 'I' },
            { '8', 'B' }
        };

        private static string NormalizeText ( string rawText )
        {
            if ( string.IsNullOrWhiteSpace( rawText ) )
            {
                return string.Empty;
            }

            var lines = rawText
                .Split( new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries )
                .Select( NormalizeLine )
                .Where( line => !string.IsNullOrWhiteSpace( line ) )
                .ToArray();

            return string.Join( Environment.NewLine, lines );
        }

        private static string NormalizeLine ( string rawLine )
        {
            if ( string.IsNullOrWhiteSpace( rawLine ) )
            {
                return string.Empty;
            }

            var filteredChars = rawLine
                .Select( char.ToUpperInvariant )
                .Where( AllowedCharacterSet.Contains )
                .ToArray();

            if ( filteredChars.Length == 0 )
            {
                return string.Empty;
            }

            var filtered = new string( filteredChars );
            var lettersOnly = new string( filtered
                .Replace( "-", string.Empty )
                .Select( c => SoftEquivalents.TryGetValue( c, out var mapped ) ? mapped : c )
                .ToArray() );

            if ( lettersOnly.Length < 32 )
            {
                return string.Empty;
            }

            var prefixLength = Math.Min( 32, lettersOnly.Length );
            var prefix = lettersOnly.Substring( 0, prefixLength );

            var remainder = lettersOnly.Length > 32
                ? lettersOnly.Substring( 32 )
                : string.Empty;

            if ( prefix.Length < 32 && remainder.Length > 0 )
            {
                var needed = Math.Min( 32 - prefix.Length, remainder.Length );
                prefix += remainder.Substring( 0, needed );
                remainder = remainder.Substring( needed );
            }

            if ( remainder.Length > 2 )
            {
                remainder = remainder.Substring( 0, 2 );
            }

            if ( remainder.Length == 0 )
            {
                return FormatPrefix( prefix );
            }

            return $"{FormatPrefix( prefix )}-{remainder}";
        }

        private static int CountValidLines ( string normalizedText )
        {
            if ( string.IsNullOrEmpty( normalizedText ) )
            {
                return 0;
            }

            return normalizedText
                .Split( new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries )
                .Count( line => line.Length >= 32 );
        }

        private static string FormatPrefix ( string prefix )
        {
            if ( string.IsNullOrEmpty( prefix ) )
            {
                return string.Empty;
            }

            var groups = new List<string>();
            for ( var i = 0; i < prefix.Length; i += 4 )
            {
                var length = Math.Min( 4, prefix.Length - i );
                groups.Add( prefix.Substring( i, length ) );
            }

            return string.Join( " ", groups );
        }

        private static string PreprocessImage ( string originalPath )
        {
            var tempPath = Path.Combine( Path.GetTempPath(), $"simpleocr-pre-{Guid.NewGuid():N}.png" );

            using ( var src = Cv2.ImRead( originalPath, ImreadModes.Color ) )
            {
                if ( src.Empty() )
                {
                    throw new InvalidOperationException( "画像を読み込めませんでした。" );
                }

                using ( var gray = new Mat() )
                using ( var blurred = new Mat() )
                using ( var binary = new Mat() )
                using ( var closed = new Mat() )
                using ( var scaled = new Mat() )
                {
                    Cv2.CvtColor( src, gray, ColorConversionCodes.BGR2GRAY );
                    Cv2.GaussianBlur( gray, blurred, new OpenCvSharp.Size( 3, 3 ), 0 );
                    Cv2.AdaptiveThreshold(
                        blurred,
                        binary,
                        maxValue: 255,
                        adaptiveMethod: AdaptiveThresholdTypes.MeanC,
                        thresholdType: ThresholdTypes.Binary,
                        blockSize: 17,
                        c: 10 );

                    using ( var kernel = Cv2.GetStructuringElement( MorphShapes.Rect, new OpenCvSharp.Size( 2, 2 ) ) )
                    {
                        Cv2.MorphologyEx( binary, closed, MorphTypes.Close, kernel );
                    }

                    Cv2.Resize( closed, scaled, new OpenCvSharp.Size(), 1.6, 1.6, InterpolationFlags.Linear );
                    Cv2.ImWrite( tempPath, scaled );
                }
            }

            return tempPath;
        }
    }
}
