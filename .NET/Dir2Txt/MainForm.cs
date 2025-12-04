using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            Load += MainForm_Load;
            FormClosed += MainForm_FormClosed;
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
                var ignoreDirs = ParseIgnoreList( txtIgnoreDirs.Text );
                var ignoreFiles = ParseIgnoreList( txtIgnoreFiles.Text );
                txtOutput.Text = BuildDirectoryText( dirPath, ignoreDirs, ignoreFiles );
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

        private string BuildDirectoryText ( string rootPath, IEnumerable<string> ignoreDirs, IEnumerable<string> ignoreFiles )
        {
            var ignoreDirSet = new HashSet<string>( ignoreDirs, StringComparer.OrdinalIgnoreCase );
            var ignoreFileSet = new HashSet<string>( ignoreFiles, StringComparer.OrdinalIgnoreCase );

            var allFiles = Directory.GetFiles( rootPath, "*", SearchOption.AllDirectories )
                                    .OrderBy( x => x )
                                    .Where( file => !ShouldIgnoreFile( file, ignoreDirSet, ignoreFileSet ) )
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

        private IEnumerable<string> ParseIgnoreList ( string text )
        {
            if ( string.IsNullOrWhiteSpace( text ) )
            {
                yield break;
            }

            var parts = text.Split( new[] { '/' }, StringSplitOptions.RemoveEmptyEntries );
            foreach ( var part in parts )
            {
                var trimmed = part.Trim();
                if ( !string.IsNullOrEmpty( trimmed ) )
                {
                    yield return trimmed;
                }
            }
        }

        private bool ShouldIgnoreFile ( string filePath, HashSet<string> ignoreDirs, HashSet<string> ignoreFiles )
        {
            if ( ignoreFiles.Contains( Path.GetFileName( filePath ) ) )
            {
                return true;
            }

            if ( ignoreDirs.Count == 0 )
            {
                return false;
            }

            var directory = Path.GetDirectoryName( filePath );
            if ( string.IsNullOrEmpty( directory ) )
            {
                return false;
            }

            var parts = directory.Split( Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar );
            return parts.Any( part => ignoreDirs.Contains( part ) );
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

        private void MainForm_Load ( object sender, EventArgs e )
        {
            try
            {
                var version = Assembly.GetExecutingAssembly().GetName().Version;
                this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";

                var settings = LoadSettings();
                if ( settings != null )
                {
                    txtDirPath.Text = settings.DirPath;
                    txtIgnoreDirs.Text = settings.IgnoreDirs;
                    txtIgnoreFiles.Text = settings.IgnoreFiles;
                }
            }
            catch
            {
            }
        }

        private void MainForm_FormClosed ( object sender, FormClosedEventArgs e )
        {
            try
            {
                SaveSettings( new MainFormSettings
                {
                    DirPath = txtDirPath.Text ?? string.Empty,
                    IgnoreDirs = txtIgnoreDirs.Text ?? string.Empty,
                    IgnoreFiles = txtIgnoreFiles.Text ?? string.Empty
                } );
            }
            catch
            {
            }
        }

        private string GetSettingsPath ()
        {
            return Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "Dir2Txt.settings.json" );
        }

        private MainFormSettings LoadSettings ()
        {
            var path = GetSettingsPath();
            if ( !File.Exists( path ) )
            {
                return null;
            }

            using ( var stream = File.OpenRead( path ) )
            {
                var serializer = new DataContractJsonSerializer( typeof( MainFormSettings ) );
                return serializer.ReadObject( stream ) as MainFormSettings;
            }
        }

        private void SaveSettings ( MainFormSettings settings )
        {
            var path = GetSettingsPath();
            using ( var stream = File.Create( path ) )
            {
                var serializer = new DataContractJsonSerializer( typeof( MainFormSettings ) );
                serializer.WriteObject( stream, settings );
            }
        }

        [DataContract]
        private class MainFormSettings
        {
            [DataMember]
            public string DirPath { get; set; }

            [DataMember]
            public string IgnoreDirs { get; set; }

            [DataMember]
            public string IgnoreFiles { get; set; }
        }


    }
}
