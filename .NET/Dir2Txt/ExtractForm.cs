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
    public partial class ExtractForm : Form
    {
        public ExtractForm () : this( string.Empty )
        {
        }

        public ExtractForm ( string initialText )
        {
            InitializeComponent();
            txtOutput.Text = initialText ?? string.Empty;
            txtExtractDirPath.AllowDrop = true;
            txtExtractDirPath.DragEnter += PathTextBox_DragEnter;
            txtExtractDirPath.DragDrop += TxtExtractDirPath_DragDrop;
            btnRun.Click += BtnRun_Click;
            btnRefExtractDirPath.Click += BtnRefExtractDirPath_Click;
        }

        private void BtnRefExtractDirPath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new FolderBrowserDialog() )
            {
                dialog.Description = "復元先フォルダを選択してください";
                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtExtractDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnRun_Click ( object sender, EventArgs e )
        {
            var targetDir = txtExtractDirPath.Text;
            if ( string.IsNullOrWhiteSpace( targetDir ) )
            {
                MessageBox.Show( this, "復元フォルダパスを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            var entries = ParseEntries( txtOutput.Text );
            if ( entries.Count == 0 )
            {
                MessageBox.Show( this, "復元対象のテキストに有効なデータがありません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( MessageBox.Show( this, "テキストの内容を復元します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question ) != DialogResult.Yes )
            {
                return;
            }

            if ( Directory.Exists( targetDir ) )
            {
                if ( MessageBox.Show( this, "復元先フォルダは既に存在します。上書きしてもよろしいですか？", "上書き確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning ) != DialogResult.Yes )
                {
                    return;
                }
            }
            else
            {
                Directory.CreateDirectory( targetDir );
            }

            var firstPath = entries.First().OriginalPath;
            var baseRoot = Path.GetPathRoot( firstPath ) ?? string.Empty;
            if ( string.IsNullOrEmpty( baseRoot ) )
            {
                baseRoot = GetCommonRootPath( entries.Select( x => x.OriginalPath ).ToList() );
            }
            if ( string.IsNullOrEmpty( baseRoot ) )
            {
                baseRoot = Path.GetDirectoryName( firstPath ) ?? firstPath;
            }

            try
            {
                foreach ( var entry in entries )
                {
                    var relative = GetRelativePath( baseRoot, entry.OriginalPath );
                    if ( string.IsNullOrEmpty( relative ) )
                    {
                        relative = Path.GetFileName( entry.OriginalPath );
                    }

                    var destination = Path.Combine( targetDir, relative );
                    var directoryName = Path.GetDirectoryName( destination );
                    if ( !string.IsNullOrEmpty( directoryName ) && !Directory.Exists( directoryName ) )
                    {
                        Directory.CreateDirectory( directoryName );
                    }

                    var encoding = GetEncodingOrDefault( entry.EncodingName );
                    using ( var writer = new StreamWriter( destination, false, encoding ) )
                    {
                        writer.Write( entry.Content );
                    }
                }

                MessageBox.Show( this, "復元が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"復元に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        private List<FileEntry> ParseEntries ( string rawText )
        {
            var result = new List<FileEntry>();
            using ( var reader = new StringReader( rawText ?? string.Empty ) )
            {
                string line;
                string currentPath = null;
                string currentEncoding = "utf-8";
                var contentBuilder = new StringBuilder();

                while ( ( line = reader.ReadLine() ) != null )
                {
                    if ( line.StartsWith( "@@" ) )
                    {
                        if ( currentPath != null )
                        {
                            result.Add( new FileEntry( currentPath, currentEncoding, contentBuilder.ToString() ) );
                            contentBuilder.Clear();
                        }

                        var meta = line.Substring( 2 ).Trim();
                        var parts = meta.Split( new[] { '|' }, 2 );
                        currentPath = parts[0];
                        currentEncoding = parts.Length > 1 && !string.IsNullOrWhiteSpace( parts[1] ) ? parts[1].Trim() : "utf-8";
                    }
                    else if ( currentPath != null )
                    {
                        contentBuilder.AppendLine( line );
                    }
                }

                if ( currentPath != null )
                {
                    result.Add( new FileEntry( currentPath, currentEncoding, contentBuilder.ToString() ) );
                }
            }

            return result;
        }

        private void PathTextBox_DragEnter ( object sender, DragEventArgs e )
        {
            e.Effect = e.Data.GetDataPresent( DataFormats.FileDrop ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void TxtExtractDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var paths = e.Data.GetData( DataFormats.FileDrop ) as string[];
            var path = paths?.FirstOrDefault();
            if ( string.IsNullOrEmpty( path ) )
            {
                return;
            }

            if ( Directory.Exists( path ) )
            {
                txtExtractDirPath.Text = path;
            }
            else if ( File.Exists( path ) )
            {
                var dir = Path.GetDirectoryName( path );
                if ( !string.IsNullOrEmpty( dir ) )
                {
                    txtExtractDirPath.Text = dir;
                }
            }
        }

        private string GetCommonRootPath ( List<string> paths )
        {
            if ( paths == null || paths.Count == 0 )
            {
                return string.Empty;
            }

            var segmentsList = paths.Select( SplitSegments ).ToList();
            var commonLength = segmentsList.Min( s => s.Count );

            for ( int i = 0; i < commonLength; i++ )
            {
                var first = segmentsList[0][i];
                if ( segmentsList.Any( s => !string.Equals( s[i], first, StringComparison.OrdinalIgnoreCase ) ) )
                {
                    commonLength = i;
                    break;
                }
            }

            if ( commonLength == 0 )
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            builder.Append( segmentsList[0][0] );
            if ( segmentsList[0][0].EndsWith( ":" ) )
            {
                builder.Append( "\\" );
            }

            for ( int i = 1; i < commonLength; i++ )
            {
                if ( builder.Length > 0 && builder[builder.Length - 1] != '\\' )
                {
                    builder.Append( "\\" );
                }

                builder.Append( segmentsList[0][i] );
            }

            return builder.ToString().TrimEnd( '\\' );
        }

        private List<string> SplitSegments ( string path )
        {
            return ( path ?? string.Empty )
                .Replace( '/', '\\' )
                .Split( new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries )
                .ToList();
        }

        private string GetRelativePath ( string basePath, string targetPath )
        {
            if ( string.IsNullOrEmpty( basePath ) || string.IsNullOrEmpty( targetPath ) )
            {
                return string.Empty;
            }

            var normalizedBase = basePath.Replace( '/', '\\' ).TrimEnd( '\\' ) + "\\";
            var normalizedTarget = targetPath.Replace( '/', '\\' );

            if ( normalizedTarget.StartsWith( normalizedBase, StringComparison.OrdinalIgnoreCase ) )
            {
                return normalizedTarget.Substring( normalizedBase.Length );
            }

            return string.Empty;
        }

        private Encoding GetEncodingOrDefault ( string name )
        {
            if ( string.IsNullOrWhiteSpace( name ) )
            {
                return new UTF8Encoding( false );
            }

            if ( string.Equals( name, "utf-8", StringComparison.OrdinalIgnoreCase ) )
            {
                return new UTF8Encoding( false );
            }

            try
            {
                return Encoding.GetEncoding( name );
            }
            catch
            {
                try
                {
                    return Encoding.GetEncoding( 932 );
                }
                catch
                {
                    return new UTF8Encoding( false );
                }
            }
        }

        private class FileEntry
        {
            public FileEntry ( string originalPath, string encodingName, string content )
            {
                OriginalPath = originalPath;
                EncodingName = encodingName;
                Content = content;
            }

            public string OriginalPath { get; }
            public string EncodingName { get; }
            public string Content { get; }
        }
    }
}
