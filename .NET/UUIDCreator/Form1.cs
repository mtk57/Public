using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace UUIDCreator
{
    public partial class Form1 : Form
    {
        private static readonly Regex CountLineRegex = new Regex( "^[0-9]+$", RegexOptions.Compiled );
        private static readonly Encoding ShiftJisEncoding = Encoding.GetEncoding( 932 );

        public Form1 ()
        {
            InitializeComponent();
            InitializeUiEvents();
        }

        private void InitializeUiEvents ()
        {
            btnRefInputFilePath.Click += BtnRefInputFilePath_Click;
            btnRefOutputDirPath.Click += BtnRefOutputDirPath_Click;
            btnCreate.Click += BtnCreate_Click;

            txtInputFilePath.AllowDrop = true;
            txtOutputDirPath.AllowDrop = true;

            txtInputFilePath.DragEnter += TxtInputFilePath_DragEnter;
            txtInputFilePath.DragDrop += TxtInputFilePath_DragDrop;
            txtOutputDirPath.DragEnter += TxtOutputDirPath_DragEnter;
            txtOutputDirPath.DragDrop += TxtOutputDirPath_DragDrop;
        }

        private void BtnRefInputFilePath_Click ( object sender, EventArgs e )
        {
            using ( OpenFileDialog dialog = new OpenFileDialog() )
            {
                dialog.Title = "件数ファイルを選択";
                dialog.Filter = "テキストファイル (*.txt)|*.txt|すべてのファイル (*.*)|*.*";
                dialog.CheckFileExists = true;

                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtInputFilePath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRefOutputDirPath_Click ( object sender, EventArgs e )
        {
            using ( FolderBrowserDialog dialog = new FolderBrowserDialog() )
            {
                dialog.Description = "出力フォルダを選択";

                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtOutputDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void TxtInputFilePath_DragEnter ( object sender, DragEventArgs e )
        {
            e.Effect = HasSinglePath( e ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void TxtInputFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            string path = GetSinglePath( e );
            if ( path != null )
            {
                txtInputFilePath.Text = path;
            }
        }

        private void TxtOutputDirPath_DragEnter ( object sender, DragEventArgs e )
        {
            e.Effect = HasSinglePath( e ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void TxtOutputDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            string path = GetSinglePath( e );
            if ( path != null )
            {
                txtOutputDirPath.Text = path;
            }
        }

        private static bool HasSinglePath ( DragEventArgs e )
        {
            if ( !e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                return false;
            }

            string[] items = ( string[] )e.Data.GetData( DataFormats.FileDrop );
            return items != null && items.Length == 1;
        }

        private static string GetSinglePath ( DragEventArgs e )
        {
            if ( !e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                return null;
            }

            string[] items = ( string[] )e.Data.GetData( DataFormats.FileDrop );
            if ( items == null || items.Length != 1 )
            {
                return null;
            }

            return items[0];
        }

        private void BtnCreate_Click ( object sender, EventArgs e )
        {
            try
            {
                string inputFilePath = txtInputFilePath.Text.Trim();
                if ( string.IsNullOrWhiteSpace( inputFilePath ) )
                {
                    ShowError( "件数ファイルパスは必須です。" );
                    return;
                }

                if ( !File.Exists( inputFilePath ) )
                {
                    ShowError( "件数ファイルが存在しません。" );
                    return;
                }

                ParsedInput parsedInput;
                try
                {
                    parsedInput = ParseAndValidateInputFile( inputFilePath );
                }
                catch ( Exception ex )
                {
                    ShowError( ex.Message );
                    return;
                }

                string outputDirPath = ResolveOutputDirectory( inputFilePath, txtOutputDirPath.Text.Trim() );
                Directory.CreateDirectory( outputDirPath );

                foreach ( int count in parsedInput.Counts )
                {
                    WriteUuidFile( outputDirPath, count );
                }

                MessageBox.Show(
                    this,
                    $"UUIDファイルを作成しました。\r\n出力先: {outputDirPath}",
                    "完了",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                ShowError( $"予期しないエラーが発生しました。\r\n{ex.Message}" );
            }
        }

        private static ParsedInput ParseAndValidateInputFile ( string inputFilePath )
        {
            byte[] bytes = File.ReadAllBytes( inputFilePath );
            ValidateNewLineCodes( bytes );

            Encoding encoding = DetectAllowedEncoding( bytes );
            string text = encoding.GetString( bytes ).TrimStart( '\uFEFF' );

            if ( text.IndexOf( '\0' ) >= 0 )
            {
                throw new InvalidOperationException( "入力ファイルのエンコードはS-JISまたはUTF-8のみ対応しています。" );
            }

            string[] lines = text.Split( new[] { "\r\n", "\n" }, StringSplitOptions.None );
            List<int> counts = new List<int>();

            for ( int i = 0; i < lines.Length; i++ )
            {
                string line = lines[i].Trim();
                if ( line.Length == 0 )
                {
                    continue;
                }

                if ( !CountLineRegex.IsMatch( line ) )
                {
                    throw new InvalidOperationException( $"{i + 1}行目が不正です。件数は半角数字のみ指定してください。" );
                }

                if ( !int.TryParse( line, out int count ) )
                {
                    throw new InvalidOperationException( $"{i + 1}行目の件数が大きすぎます。" );
                }

                counts.Add( count );
            }

            if ( counts.Count == 0 )
            {
                throw new InvalidOperationException( "入力ファイルに有効な件数がありません。" );
            }

            return new ParsedInput( counts );
        }

        private static void ValidateNewLineCodes ( byte[] bytes )
        {
            for ( int i = 0; i < bytes.Length; i++ )
            {
                if ( bytes[i] != 0x0D )
                {
                    continue;
                }

                if ( i + 1 >= bytes.Length || bytes[i + 1] != 0x0A )
                {
                    throw new InvalidOperationException( "改行コードはCRLFまたはLFのみ使用できます。" );
                }
            }
        }

        private static Encoding DetectAllowedEncoding ( byte[] bytes )
        {
            if ( HasUtf8Bom( bytes ) )
            {
                return new UTF8Encoding( true, true );
            }

            bool utf8Valid = IsValidEncoding( bytes, new UTF8Encoding( false, true ) );
            bool shiftJisValid = IsValidEncoding( bytes, Encoding.GetEncoding( 932, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback ) );

            if ( utf8Valid )
            {
                return new UTF8Encoding( false, true );
            }

            if ( shiftJisValid )
            {
                return ShiftJisEncoding;
            }

            throw new InvalidOperationException( "入力ファイルのエンコードはS-JISまたはUTF-8のみ対応しています。" );
        }

        private static bool IsValidEncoding ( byte[] bytes, Encoding encoding )
        {
            try
            {
                encoding.GetString( bytes );
                return true;
            }
            catch ( DecoderFallbackException )
            {
                return false;
            }
        }

        private static bool HasUtf8Bom ( byte[] bytes )
        {
            return bytes.Length >= 3
                && bytes[0] == 0xEF
                && bytes[1] == 0xBB
                && bytes[2] == 0xBF;
        }

        private static string ResolveOutputDirectory ( string inputFilePath, string outputDirPath )
        {
            if ( !string.IsNullOrWhiteSpace( outputDirPath ) && Directory.Exists( outputDirPath ) )
            {
                return outputDirPath;
            }

            string inputDirectory = Path.GetDirectoryName( inputFilePath ) ?? AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine( inputDirectory, "UUID" );
        }

        private static void WriteUuidFile ( string outputDirPath, int count )
        {
            string basePath = Path.Combine( outputDirPath, $"{count}.txt" );
            string filePath = GetUniqueOutputPath( basePath );

            List<string> lines = new List<string>( count );
            for ( int i = 0; i < count; i++ )
            {
                lines.Add( Guid.NewGuid().ToString() );
            }

            string content = string.Join( "\r\n", lines );
            File.WriteAllText( filePath, content, ShiftJisEncoding );
        }

        private static string GetUniqueOutputPath ( string basePath )
        {
            if ( !File.Exists( basePath ) )
            {
                return basePath;
            }

            string directory = Path.GetDirectoryName( basePath ) ?? string.Empty;
            string nameWithoutExtension = Path.GetFileNameWithoutExtension( basePath );
            string extension = Path.GetExtension( basePath );
            string timestamp = DateTime.Now.ToString( "yyyyMMdd_HHmmssfff" );
            string candidate = Path.Combine( directory, $"{nameWithoutExtension}_{timestamp}{extension}" );

            if ( !File.Exists( candidate ) )
            {
                return candidate;
            }

            int suffix = 1;
            while ( true )
            {
                string fallback = Path.Combine( directory, $"{nameWithoutExtension}_{timestamp}_{suffix}{extension}" );
                if ( !File.Exists( fallback ) )
                {
                    return fallback;
                }

                suffix++;
            }
        }

        private void ShowError ( string message )
        {
            MessageBox.Show( this, message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
        }

        private sealed class ParsedInput
        {
            public ParsedInput ( List<int> counts )
            {
                Counts = counts;
            }

            public List<int> Counts { get; }
        }
    }
}
