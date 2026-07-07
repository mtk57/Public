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
using System.Diagnostics;
using System.Runtime.Serialization.Json;

namespace Dir2Txt
{
    public partial class ExtractForm : Form
    {
        private const int MaxExtractDirPathHistoryCount = 20;

        public ExtractForm () : this( string.Empty )
        {
        }

        public ExtractForm ( string initialText )
        {
            InitializeComponent();
            txtOutput.Text = initialText ?? string.Empty;
            cmbExtractDirPath.AllowDrop = true;
            cmbExtractDirPath.DragEnter += PathTextBox_DragEnter;
            cmbExtractDirPath.DragDrop += CmbExtractDirPath_DragDrop;
            cmbExtractDirPath.Leave += CmbExtractDirPath_Leave;
            cmbExtractDirPath.KeyDown += CmbExtractDirPath_KeyDown;
            btnRun.Click += BtnRun_Click;
            btnRefExtractDirPath.Click += BtnRefExtractDirPath_Click;
            btnClearExtractDirPath.Click += BtnClearExtractDirPath_Click;
            LoadExtractDirPathHistories();
        }

        private void BtnRefExtractDirPath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new FolderBrowserDialog() )
            {
                dialog.Description = "復元先フォルダを選択してください";
                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    cmbExtractDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnRun_Click ( object sender, EventArgs e )
        {
            var targetDir = ( cmbExtractDirPath.Text ?? string.Empty ).Trim();
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
                        writer.Write( NormalizeLineEndings( entry.Content, GetLineEndingText( entry.LineEndingName ) ) );
                    }
                }

                MessageBox.Show( this, "復元が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information );
                OpenFolder( targetDir );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"復元に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        private void BtnClearExtractDirPath_Click ( object sender, EventArgs e )
        {
            cmbExtractDirPath.Text = string.Empty;
        }

        private void CmbExtractDirPath_Leave ( object sender, EventArgs e )
        {
            AddExtractDirPathHistory();
        }

        private void CmbExtractDirPath_KeyDown ( object sender, KeyEventArgs e )
        {
            if ( e.KeyCode != Keys.Enter )
            {
                return;
            }

            AddExtractDirPathHistory();
            e.SuppressKeyPress = true;
        }

        private void AddExtractDirPathHistory ()
        {
            var path = ( cmbExtractDirPath.Text ?? string.Empty ).Trim();
            if ( string.IsNullOrEmpty( path ) )
            {
                return;
            }

            var histories = cmbExtractDirPath.Items.Cast<string>()
                .Where( x => !string.IsNullOrWhiteSpace( x ) )
                .Where( x => !string.Equals( x, path, StringComparison.OrdinalIgnoreCase ) )
                .ToList();

            histories.Insert( 0, path );
            histories = histories.Take( MaxExtractDirPathHistoryCount ).ToList();
            SetExtractDirPathHistories( histories );
            SaveExtractDirPathHistories( histories );
        }

        private void LoadExtractDirPathHistories ()
        {
            var settings = LoadSettings();
            var histories = settings?.ExtractDirPathHistories ?? new List<string>();
            SetExtractDirPathHistories( histories );
        }

        private void SetExtractDirPathHistories ( IEnumerable<string> histories )
        {
            var currentText = cmbExtractDirPath.Text;
            var items = histories
                .Where( x => !string.IsNullOrWhiteSpace( x ) )
                .Select( x => x.Trim() )
                .Distinct( StringComparer.OrdinalIgnoreCase )
                .Take( MaxExtractDirPathHistoryCount )
                .ToList();

            cmbExtractDirPath.Items.Clear();
            cmbExtractDirPath.Items.AddRange( items.Cast<object>().ToArray() );
            if ( string.IsNullOrWhiteSpace( currentText ) && items.Count > 0 )
            {
                cmbExtractDirPath.SelectedIndex = 0;
                return;
            }

            cmbExtractDirPath.Text = currentText;
        }

        private void SaveExtractDirPathHistories ( List<string> histories )
        {
            var settings = LoadSettings() ?? new AppSettings();
            settings.ExtractDirPathHistories = histories;
            SaveSettings( settings );
        }

        private string GetSettingsPath ()
        {
            return Path.Combine( AppDomain.CurrentDomain.BaseDirectory, "Dir2Txt.settings.json" );
        }

        private AppSettings LoadSettings ()
        {
            var path = GetSettingsPath();
            if ( !File.Exists( path ) )
            {
                return null;
            }

            using ( var stream = File.OpenRead( path ) )
            {
                var serializer = new DataContractJsonSerializer( typeof( AppSettings ) );
                return serializer.ReadObject( stream ) as AppSettings;
            }
        }

        private void SaveSettings ( AppSettings settings )
        {
            var path = GetSettingsPath();
            using ( var stream = File.Create( path ) )
            {
                var serializer = new DataContractJsonSerializer( typeof( AppSettings ) );
                serializer.WriteObject( stream, settings );
            }
        }

        private void OpenFolder ( string folderPath )
        {
            Process.Start( "explorer.exe", "\"" + folderPath + "\"" );
        }

        private List<FileEntry> ParseEntries ( string rawText )
        {
            var parsed = ParseEntriesCore( rawText, true );
            if ( parsed.Count > 0 )
            {
                return parsed;
            }

            return ParseEntriesCore( rawText, false );
        }

        private List<FileEntry> ParseEntriesCore ( string rawText, bool skipHeader )
        {
            var result = new List<FileEntry>();
            using ( var reader = new StringReader( rawText ?? string.Empty ) )
            {
                string line;
                string currentPath = null;
                string currentEncoding = "utf-8";
                string currentLineEnding = "CRLF";
                var contentBuilder = new StringBuilder();
                var started = !skipHeader;

                while ( ( line = reader.ReadLine() ) != null )
                {
                    if ( !started )
                    {
                        if ( line == "==========" )
                        {
                            started = true;
                        }
                        continue;
                    }

                    if ( line.StartsWith( "@@" ) )
                    {
                        if ( currentPath != null )
                        {
                            result.Add( new FileEntry( currentPath, currentEncoding, currentLineEnding, contentBuilder.ToString() ) );
                            contentBuilder.Clear();
                        }

                        var meta = line.Substring( 2 ).Trim();
                        var parts = meta.Split( new[] { '|' }, 3 );
                        currentPath = parts[0];
                        currentEncoding = parts.Length > 1 && !string.IsNullOrWhiteSpace( parts[1] ) ? parts[1].Trim() : "utf-8";
                        currentLineEnding = parts.Length > 2 && !string.IsNullOrWhiteSpace( parts[2] ) ? parts[2].Trim() : "CRLF";
                    }
                    else if ( currentPath != null )
                    {
                        contentBuilder.AppendLine( line );
                    }
                }

                if ( currentPath != null )
                {
                    result.Add( new FileEntry( currentPath, currentEncoding, currentLineEnding, contentBuilder.ToString() ) );
                }
            }

            return result;
        }

        private string GetLineEndingText ( string name )
        {
            return string.Equals( name, "LF", StringComparison.OrdinalIgnoreCase ) ? "\n" : "\r\n";
        }

        private string NormalizeLineEndings ( string text, string lineEnding )
        {
            if ( string.IsNullOrEmpty( text ) )
            {
                return text;
            }

            return text.Replace( "\r\n", "\n" ).Replace( "\r", "\n" ).Replace( "\n", lineEnding );
        }

        private void PathTextBox_DragEnter ( object sender, DragEventArgs e )
        {
            e.Effect = e.Data.GetDataPresent( DataFormats.FileDrop ) ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void CmbExtractDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var paths = e.Data.GetData( DataFormats.FileDrop ) as string[];
            var path = paths?.FirstOrDefault();
            if ( string.IsNullOrEmpty( path ) )
            {
                return;
            }

            if ( Directory.Exists( path ) )
            {
                cmbExtractDirPath.Text = path;
            }
            else if ( File.Exists( path ) )
            {
                var dir = Path.GetDirectoryName( path );
                if ( !string.IsNullOrEmpty( dir ) )
                {
                    cmbExtractDirPath.Text = dir;
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
            public FileEntry ( string originalPath, string encodingName, string lineEndingName, string content )
            {
                OriginalPath = originalPath;
                EncodingName = encodingName;
                LineEndingName = lineEndingName;
                Content = content;
            }

            public string OriginalPath { get; }
            public string EncodingName { get; }
            public string LineEndingName { get; }
            public string Content { get; }
        }
    }
}
