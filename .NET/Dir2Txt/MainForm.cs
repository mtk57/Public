using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dir2Txt
{
    public partial class MainForm : Form
    {
        private const string DELIMITER = "==========";
        private static readonly Encoding OutputFileEncoding = new UTF8Encoding( true );
        private CancellationTokenSource buildCancellation;

        public MainForm ()
        {
            InitializeComponent();
            txtDirPath.AllowDrop = true;
            txtDirPath.DragEnter += PathTextBox_DragEnter;
            txtDirPath.DragDrop += TxtDirPath_DragDrop;
            btnRefDirPath.Click += BtnRefDirPath_Click;
            btnRun.Click += BtnRun_Click;
            btnCancel.Click += BtnCancel_Click;
            btnExtract.Click += BtnExtract_Click;
            btnDivide.Click += BtnDivide_Click;
            btnHelp.Click += BtnHelp_Click;
            txtOutput.TextChanged += TxtOutput_TextChanged;
            Load += MainForm_Load;
            FormClosed += MainForm_FormClosed;
            SetBuildRunning( false );
            UpdateOutputLength();
        }

        private void BtnHelp_Click ( object sender, EventArgs e )
        {
            ShowFormatHelpDialog();
        }

        private void ShowFormatHelpDialog ()
        {
            var helpText = GetFormatHelpText();
            using ( var form = new Form() )
            using ( var textBox = new TextBox() )
            using ( var copyButton = new Button() )
            using ( var closeButton = new Button() )
            using ( var buttonPanel = new FlowLayoutPanel() )
            {
                form.Text = "Dir2Txt形式の説明";
                form.StartPosition = FormStartPosition.CenterParent;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ShowIcon = false;
                form.Size = new Size( 640, 520 );

                textBox.Dock = DockStyle.Fill;
                textBox.Multiline = true;
                textBox.ReadOnly = true;
                textBox.ScrollBars = ScrollBars.Both;
                textBox.WordWrap = false;
                textBox.Text = helpText;

                copyButton.Text = "コピー";
                copyButton.AutoSize = true;
                copyButton.Click += ( s, e ) =>
                {
                    Clipboard.SetText( helpText );
                    MessageBox.Show( form, "説明文をコピーしました。", "コピー完了", MessageBoxButtons.OK, MessageBoxIcon.Information );
                };

                closeButton.Text = "閉じる";
                closeButton.AutoSize = true;
                closeButton.Click += ( s, e ) => form.Close();

                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.FlowDirection = FlowDirection.RightToLeft;
                buttonPanel.Height = 40;
                buttonPanel.Padding = new Padding( 8 );
                buttonPanel.Controls.Add( closeButton );
                buttonPanel.Controls.Add( copyButton );

                form.Controls.Add( textBox );
                form.Controls.Add( buttonPanel );
                form.AcceptButton = copyButton;
                form.CancelButton = closeButton;
                form.ShowDialog( this );
            }
        }

        private string GetFormatHelpText ()
        {
            var nl = Environment.NewLine;
            return string.Join( nl, new[]
            {
                "Dir2Txt形式の説明",
                "",
                "出力テキストは、複数ファイルを1つのテキストにまとめた独自形式です。",
                "",
                "全体構造",
                "1. 先頭から「==========」の行までは、対象ファイルのパス一覧です。",
                "2. 「==========」だけの行の次から、各ファイルの本文データが並びます。",
                "3. 「復元」は「==========」より前を読み飛ばします。「==========」がない場合は先頭から本文データとして読みます。",
                "",
                "ファイル開始行",
                "各ファイルは、次の形式の行で始まります。",
                "@@<元ファイルパス>|<文字コード>|<改行種別>",
                "",
                "例:",
                "@@C:\\work\\sample\\memo.txt|utf-8|CRLF",
                "",
                "意味:",
                "- 「@@」で始まる行は新しいファイルの開始を表します。",
                "- <元ファイルパス> はファイルの絶対パスです。",
                "- <文字コード> は保存文字コードです。例: utf-8, shift_jis",
                "- <改行種別> は改行コードです。CRLF または LF を使います。",
                "",
                "本文",
                "ファイル開始行の次の行から、次の「@@」行の直前までがそのファイルの本文です。",
                "本文を編集する場合は、通常は本文だけを変更し、「@@」行、「==========」行、パス、文字コード、改行種別は変更しないでください。",
                "",
                "注意",
                "- 行頭が「@@」の行はファイル開始行として扱われます。",
                "",
                "最小例",
                "C:\\work\\a.txt",
                "C:\\work\\b.txt",
                "==========",
                "@@C:\\work\\a.txt|utf-8|CRLF",
                "a.txt の本文",
                "@@C:\\work\\b.txt|utf-8|LF",
                "b.txt の本文"
            } );
        }

        private void BtnDivide_Click ( object sender, EventArgs e )
        {
            if ( string.IsNullOrEmpty( txtOutput.Text ) )
            {
                MessageBox.Show( this, "出力テキストがありません。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            int divideLength;
            if ( !int.TryParse( txtDivideLnegth.Text, out divideLength ) || divideLength <= 10 )
            {
                MessageBox.Show( this, "分割文字数には10より大きい数値を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            using ( var form = new DivideForm( txtOutput.Text, divideLength ) )
            {
                form.ShowDialog( this );
            }
        }

        private void BtnExtract_Click ( object sender, EventArgs e )
        {
            using ( var form = new ExtractForm( txtOutput.Text ) )
            {
                form.ShowDialog( this );
            }
        }

        private async void BtnRun_Click ( object sender, EventArgs e )
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
                var ignoreExts = ParseIgnoreList( txtIgnoreExt.Text ).ToList();
                var ignoreExtNegated = chkIgnoreExtNegated.Checked;
                var outputToFile = chkOutputToFile.Checked;
                if ( ignoreExtNegated && !ignoreExts.Any() )
                {
                    MessageBox.Show( this, "否定指定時は対象にする拡張子を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                    return;
                }

                buildCancellation = new CancellationTokenSource();
                SetBuildRunning( true );
                txtOutput.Text = string.Empty;

                var token = buildCancellation.Token;
                var progress = new Progress<BuildProgress>( UpdateBuildProgress );
                var result = await Task.Run( () =>
                    outputToFile
                        ? BuildDirectoryFile( dirPath, ignoreDirs, ignoreFiles, ignoreExts, ignoreExtNegated, progress, token )
                        : BuildDirectoryText( dirPath, ignoreDirs, ignoreFiles, ignoreExts, ignoreExtNegated, progress, token ) );
                if ( token.IsCancellationRequested || result == null )
                {
                    lblProgress.Text = "中止しました";
                    return;
                }

                if ( outputToFile )
                {
                    txtOutput.Clear();
                    lblProgress.Text = $"完了 対象: {result.TargetCount} / 出力: {result.OutputCount} / スキップ: {result.SkippedCount} / 出力先: {result.OutputPath}";
                    OpenOutputFolder( result.OutputPath );
                }
                else
                {
                    await PopulateOutputAsync( result );
                    lblProgress.Text = $"完了 対象: {result.TargetCount} / 出力: {result.OutputCount} / スキップ: {result.SkippedCount}";
                }
            }
            catch ( OperationCanceledException )
            {
                lblProgress.Text = "中止しました";
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"テキスト化に失敗しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
            finally
            {
                buildCancellation?.Dispose();
                buildCancellation = null;
                SetBuildRunning( false );
            }
        }

        private BuildResult BuildDirectoryFile ( string rootPath, IEnumerable<string> ignoreDirs, IEnumerable<string> ignoreFiles, IEnumerable<string> ignoreExts, bool ignoreExtNegated, IProgress<BuildProgress> progress, CancellationToken cancellationToken )
        {
            var ignoreDirSet = new HashSet<string>( ignoreDirs, StringComparer.OrdinalIgnoreCase );
            var ignoreFileSet = new HashSet<string>( ignoreFiles, StringComparer.OrdinalIgnoreCase );
            var ignoreExtSet = new HashSet<string>(
                ignoreExts.Select( NormalizeExtension ).Where( x => !string.IsNullOrEmpty( x ) ),
                StringComparer.OrdinalIgnoreCase );

            progress?.Report( BuildProgress.Indeterminate( "ファイル一覧を取得中..." ) );
            var targetFiles = EnumerateTargetFiles( rootPath, ignoreDirSet, ignoreFileSet, ignoreExtSet, ignoreExtNegated, cancellationToken )
                                     .OrderBy( x => x )
                                     .ToList();
            if ( cancellationToken.IsCancellationRequested )
            {
                return null;
            }

            var outputPath = CreateOutputPath( rootPath );
            var tempPath = outputPath + ".tmp";
            var processedFiles = new List<string>();
            var skippedCount = 0;
            var processedCount = 0;

            progress?.Report( BuildProgress.Determinate( 0, targetFiles.Count, $"ファイル出力中: 0 / {targetFiles.Count}" ) );
            try
            {
                using ( var tempWriter = new StreamWriter( tempPath, false, new UTF8Encoding( false ) ) )
                {
                    foreach ( var file in targetFiles )
                    {
                        if ( cancellationToken.IsCancellationRequested )
                        {
                            return null;
                        }

                        try
                        {
                            if ( IsBinaryFile( file, cancellationToken ) )
                            {
                                skippedCount++;
                                continue;
                            }

                            var read = ReadFileWithEncoding( file, cancellationToken );
                            if ( cancellationToken.IsCancellationRequested )
                            {
                                return null;
                            }

                            processedFiles.Add( file );
                            tempWriter.Write( "@@" );
                            tempWriter.Write( file );
                            tempWriter.Write( "|" );
                            tempWriter.Write( read.Encoding.WebName );
                            tempWriter.Write( "|" );
                            tempWriter.WriteLine( read.LineEnding );
                            var endsWithLineEnding = WriteNormalizedLineEndings( tempWriter, read.Content, Environment.NewLine );
                            if ( !endsWithLineEnding )
                            {
                                tempWriter.WriteLine();
                            }
                        }
                        catch ( IOException )
                        {
                            skippedCount++;
                        }
                        catch ( UnauthorizedAccessException )
                        {
                            skippedCount++;
                        }
                        catch ( NotSupportedException )
                        {
                            skippedCount++;
                        }
                        finally
                        {
                            processedCount++;
                            progress?.Report( BuildProgress.Determinate( processedCount, targetFiles.Count, $"ファイル出力中: {processedCount} / {targetFiles.Count}" ) );
                        }
                    }
                }

                if ( cancellationToken.IsCancellationRequested )
                {
                    return null;
                }

                progress?.Report( BuildProgress.Indeterminate( "出力ファイル作成中..." ) );
                using ( var writer = new StreamWriter( outputPath, false, OutputFileEncoding ) )
                {
                    foreach ( var file in processedFiles )
                    {
                        writer.WriteLine( file );
                    }
                    writer.WriteLine( DELIMITER );

                    using ( var reader = new StreamReader( tempPath, Encoding.UTF8 ) )
                    {
                        var buffer = new char[32768];
                        int read;
                        while ( ( read = reader.Read( buffer, 0, buffer.Length ) ) > 0 )
                        {
                            if ( cancellationToken.IsCancellationRequested )
                            {
                                return null;
                            }

                            writer.Write( buffer, 0, read );
                        }
                    }
                }

                return new BuildResult
                {
                    OutputPath = outputPath,
                    TargetCount = targetFiles.Count,
                    OutputCount = processedFiles.Count,
                    SkippedCount = skippedCount
                };
            }
            finally
            {
                try
                {
                    if ( File.Exists( tempPath ) )
                    {
                        File.Delete( tempPath );
                    }
                }
                catch
                {
                }
            }
        }

        private string CreateOutputPath ( string rootPath )
        {
            var fileName = "Dir2Txt_" + DateTime.Now.ToString( "yyyyMMdd_HHmmss" ) + ".txt";
            var outputDir = AppDomain.CurrentDomain.BaseDirectory;
            var outputPath = Path.Combine( outputDir, fileName );
            var index = 1;
            while ( File.Exists( outputPath ) )
            {
                outputPath = Path.Combine( outputDir, Path.GetFileNameWithoutExtension( fileName ) + "_" + index.ToString() + ".txt" );
                index++;
            }

            return outputPath;
        }

        private void OpenOutputFolder ( string outputPath )
        {
            try
            {
                var directory = Path.GetDirectoryName( outputPath );
                if ( !string.IsNullOrEmpty( directory ) && Directory.Exists( directory ) )
                {
                    Process.Start( "explorer.exe", directory );
                }
            }
            catch
            {
            }
        }

        private void BtnCancel_Click ( object sender, EventArgs e )
        {
            buildCancellation?.Cancel();
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

        private BuildResult BuildDirectoryText ( string rootPath, IEnumerable<string> ignoreDirs, IEnumerable<string> ignoreFiles, IEnumerable<string> ignoreExts, bool ignoreExtNegated, IProgress<BuildProgress> progress, CancellationToken cancellationToken )
        {
            var ignoreDirSet = new HashSet<string>( ignoreDirs, StringComparer.OrdinalIgnoreCase );
            var ignoreFileSet = new HashSet<string>( ignoreFiles, StringComparer.OrdinalIgnoreCase );
            var ignoreExtSet = new HashSet<string>(
                ignoreExts.Select( NormalizeExtension ).Where( x => !string.IsNullOrEmpty( x ) ),
                StringComparer.OrdinalIgnoreCase );

            progress?.Report( BuildProgress.Indeterminate( "ファイル一覧を取得中..." ) );
            var targetFiles = EnumerateTargetFiles( rootPath, ignoreDirSet, ignoreFileSet, ignoreExtSet, ignoreExtNegated, cancellationToken )
                                     .OrderBy( x => x )
                                     .ToList();
            if ( cancellationToken.IsCancellationRequested )
            {
                return null;
            }

            progress?.Report( BuildProgress.Determinate( 0, targetFiles.Count, $"処理中: 0 / {targetFiles.Count}" ) );

            var processedCount = 0;
            var skippedCount = 0;
            var results = new ConcurrentBag<FileReadResult>();
            var options = new ParallelOptions
            {
                MaxDegreeOfParallelism = Environment.ProcessorCount
            };

            Parallel.ForEach( targetFiles, options, ( file, state ) =>
            {
                if ( cancellationToken.IsCancellationRequested )
                {
                    state.Stop();
                    return;
                }

                try
                {
                    if ( !IsBinaryFile( file, cancellationToken ) && !cancellationToken.IsCancellationRequested )
                    {
                        try
                        {
                            var read = ReadFileWithEncoding( file, cancellationToken );
                            if ( !cancellationToken.IsCancellationRequested )
                            {
                                results.Add( new FileReadResult
                                {
                                    FilePath = file,
                                    Content = read.Content,
                                    EncodingName = read.Encoding.WebName,
                                    LineEnding = read.LineEnding
                                } );
                            }
                        }
                        catch ( IOException )
                        {
                            Interlocked.Increment( ref skippedCount );
                        }
                        catch ( UnauthorizedAccessException )
                        {
                            Interlocked.Increment( ref skippedCount );
                        }
                        catch ( NotSupportedException )
                        {
                            Interlocked.Increment( ref skippedCount );
                        }
                    }
                    else
                    {
                        Interlocked.Increment( ref skippedCount );
                    }
                }
                finally
                {
                    var current = Interlocked.Increment( ref processedCount );
                    progress?.Report( BuildProgress.Determinate( current, targetFiles.Count, $"処理中: {current} / {targetFiles.Count}" ) );
                }
            } );

            if ( cancellationToken.IsCancellationRequested )
            {
                return null;
            }

            progress?.Report( BuildProgress.Indeterminate( "出力データ作成中..." ) );
            var orderedResults = results.OrderBy( x => x.FilePath ).ToList();

            return new BuildResult
            {
                Files = orderedResults,
                TargetCount = targetFiles.Count,
                OutputCount = orderedResults.Count,
                SkippedCount = skippedCount
            };
        }

        private async Task PopulateOutputAsync ( BuildResult result )
        {
            lblProgress.Text = "出力テキスト表示中...";
            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = Math.Max( 1, result.OutputCount );
            progressBar1.Value = 0;

            txtOutput.Clear();
            foreach ( var file in result.Files )
            {
                txtOutput.AppendText( file.FilePath + Environment.NewLine );
            }
            txtOutput.AppendText( DELIMITER + Environment.NewLine );

            var current = 0;
            foreach ( var file in result.Files )
            {
                txtOutput.AppendText( "@@" + file.FilePath + "|" + file.EncodingName + "|" + file.LineEnding + Environment.NewLine );
                var endsWithLineEnding = AppendNormalizedLineEndings( txtOutput, file.Content, Environment.NewLine );
                file.Content = null;
                if ( !endsWithLineEnding )
                {
                    txtOutput.AppendText( Environment.NewLine );
                }

                current++;
                progressBar1.Value = Math.Min( current, progressBar1.Maximum );
                lblProgress.Text = $"出力テキスト表示中: {current} / {result.OutputCount}";
                if ( current % 10 == 0 )
                {
                    await Task.Yield();
                }
            }
        }

        private IEnumerable<string> EnumerateTargetFiles ( string rootPath, HashSet<string> ignoreDirs, HashSet<string> ignoreFiles, HashSet<string> ignoreExts, bool ignoreExtNegated, CancellationToken cancellationToken )
        {
            foreach ( var file in EnumerateFilesSafely( rootPath, ignoreDirs, cancellationToken ) )
            {
                if ( cancellationToken.IsCancellationRequested )
                {
                    yield break;
                }

                if ( !ShouldIgnoreFile( file, ignoreDirs, ignoreFiles, ignoreExts, ignoreExtNegated ) )
                {
                    yield return file;
                }
            }
        }

        private IEnumerable<string> EnumerateFilesSafely ( string directory, HashSet<string> ignoreDirs, CancellationToken cancellationToken )
        {
            if ( cancellationToken.IsCancellationRequested )
            {
                yield break;
            }

            var dirName = Path.GetFileName( directory.TrimEnd( Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar ) );
            if ( !string.IsNullOrEmpty( dirName ) && ignoreDirs.Contains( dirName ) )
            {
                yield break;
            }

            string[] files;
            try
            {
                files = Directory.GetFiles( directory );
            }
            catch
            {
                yield break;
            }

            foreach ( var file in files )
            {
                if ( cancellationToken.IsCancellationRequested )
                {
                    yield break;
                }

                yield return file;
            }

            string[] directories;
            try
            {
                directories = Directory.GetDirectories( directory );
            }
            catch
            {
                yield break;
            }

            foreach ( var childDirectory in directories )
            {
                if ( cancellationToken.IsCancellationRequested )
                {
                    yield break;
                }

                foreach ( var file in EnumerateFilesSafely( childDirectory, ignoreDirs, cancellationToken ) )
                {
                    yield return file;
                }
            }
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

        private string NormalizeExtension ( string ext )
        {
            if ( string.IsNullOrWhiteSpace( ext ) )
            {
                return string.Empty;
            }

            var trimmed = ext.Trim();
            return trimmed.StartsWith( "." ) ? trimmed.Substring( 1 ) : trimmed;
        }

        private bool ShouldIgnoreFile ( string filePath, HashSet<string> ignoreDirs, HashSet<string> ignoreFiles, HashSet<string> ignoreExts, bool ignoreExtNegated )
        {
            var fileName = Path.GetFileName( filePath );
            if ( ignoreFiles.Contains( fileName ) )
            {
                return true;
            }

            if ( ignoreExts.Count > 0 || ignoreExtNegated )
            {
                var ext = NormalizeExtension( Path.GetExtension( filePath ) );
                var containsExt = !string.IsNullOrEmpty( ext ) && ignoreExts.Contains( ext );
                if ( ignoreExtNegated ? !containsExt : containsExt )
                {
                    return true;
                }
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

        private (string Content, Encoding Encoding, string LineEnding) ReadFileWithEncoding ( string path, CancellationToken cancellationToken )
        {
            var bytes = File.ReadAllBytes( path );
            if ( cancellationToken.IsCancellationRequested )
            {
                return ( string.Empty, Encoding.Default, "CRLF" );
            }

            var detected = DetectEncoding( bytes );
            var content = detected.GetString( bytes );
            return ( content, detected, DetectLineEnding( content ) );
        }

        private string DetectLineEnding ( string content )
        {
            if ( string.IsNullOrEmpty( content ) )
            {
                return "CRLF";
            }

            var crlfIndex = content.IndexOf( "\r\n", StringComparison.Ordinal );
            if ( crlfIndex >= 0 )
            {
                return "CRLF";
            }

            return content.IndexOf( "\n", StringComparison.Ordinal ) >= 0 ? "LF" : "CRLF";
        }

        private bool AppendNormalizedLineEndings ( TextBox textBox, string text, string lineEnding )
        {
            if ( string.IsNullOrEmpty( text ) )
            {
                return false;
            }

            const int chunkSize = 32768;
            var chunk = new StringBuilder( chunkSize );
            var endsWithLineEnding = false;
            for ( int i = 0; i < text.Length; i++ )
            {
                var current = text[i];
                if ( current == '\r' )
                {
                    if ( i + 1 < text.Length && text[i + 1] == '\n' )
                    {
                        i++;
                    }

                    chunk.Append( lineEnding );
                    endsWithLineEnding = true;
                }
                else if ( current == '\n' )
                {
                    chunk.Append( lineEnding );
                    endsWithLineEnding = true;
                }
                else
                {
                    chunk.Append( current );
                    endsWithLineEnding = false;
                }

                if ( chunk.Length >= chunkSize )
                {
                    textBox.AppendText( chunk.ToString() );
                    chunk.Clear();
                }
            }

            if ( chunk.Length > 0 )
            {
                textBox.AppendText( chunk.ToString() );
            }

            return endsWithLineEnding;
        }

        private bool WriteNormalizedLineEndings ( TextWriter writer, string text, string lineEnding )
        {
            if ( string.IsNullOrEmpty( text ) )
            {
                return false;
            }

            const int chunkSize = 32768;
            var chunk = new StringBuilder( chunkSize );
            var endsWithLineEnding = false;
            for ( int i = 0; i < text.Length; i++ )
            {
                var current = text[i];
                if ( current == '\r' )
                {
                    if ( i + 1 < text.Length && text[i + 1] == '\n' )
                    {
                        i++;
                    }

                    chunk.Append( lineEnding );
                    endsWithLineEnding = true;
                }
                else if ( current == '\n' )
                {
                    chunk.Append( lineEnding );
                    endsWithLineEnding = true;
                }
                else
                {
                    chunk.Append( current );
                    endsWithLineEnding = false;
                }

                if ( chunk.Length >= chunkSize )
                {
                    writer.Write( chunk.ToString() );
                    chunk.Clear();
                }
            }

            if ( chunk.Length > 0 )
            {
                writer.Write( chunk.ToString() );
            }

            return endsWithLineEnding;
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

        private bool IsBinaryFile ( string path, CancellationToken cancellationToken )
        {
            const int sampleSize = 8000;
            var buffer = new byte[sampleSize];
            try
            {
                using ( var stream = File.OpenRead( path ) )
                {
                    var read = stream.Read( buffer, 0, buffer.Length );
                    if ( HasTextBom( buffer, read ) )
                    {
                        return false;
                    }

                    for ( int i = 0; i < read; i++ )
                    {
                        if ( cancellationToken.IsCancellationRequested )
                        {
                            return true;
                        }

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

        private bool HasTextBom ( byte[] bytes, int length )
        {
            if ( length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF )
            {
                return true;
            }

            if ( length >= 2 )
            {
                if ( bytes[0] == 0xFF && bytes[1] == 0xFE )
                {
                    return true;
                }

                if ( bytes[0] == 0xFE && bytes[1] == 0xFF )
                {
                    return true;
                }
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
                    txtDirPath.Text = settings.DirPath ?? string.Empty;
                    txtIgnoreDirs.Text = settings.IgnoreDirs ?? string.Empty;
                    txtIgnoreFiles.Text = settings.IgnoreFiles ?? string.Empty;
                    txtIgnoreExt.Text = settings.IgnoreExt ?? string.Empty;
                    chkIgnoreExtNegated.Checked = settings.IgnoreExtNegated;
                    chkOutputToFile.Checked = settings.OutputToFile;
                    txtDivideLnegth.Text = settings.DivideLength ?? string.Empty;
                }
            }
            catch
            {
            }
        }

        private void TxtOutput_TextChanged ( object sender, EventArgs e )
        {
            UpdateOutputLength();
        }

        private void MainForm_FormClosed ( object sender, FormClosedEventArgs e )
        {
            try
            {
                var settings = LoadSettings() ?? new AppSettings();
                settings.DirPath = txtDirPath.Text ?? string.Empty;
                settings.IgnoreDirs = txtIgnoreDirs.Text ?? string.Empty;
                settings.IgnoreFiles = txtIgnoreFiles.Text ?? string.Empty;
                settings.IgnoreExt = txtIgnoreExt.Text ?? string.Empty;
                settings.IgnoreExtNegated = chkIgnoreExtNegated.Checked;
                settings.OutputToFile = chkOutputToFile.Checked;
                settings.DivideLength = txtDivideLnegth.Text ?? string.Empty;

                SaveSettings( settings );
            }
            catch
            {
            }
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

        private void UpdateOutputLength ()
        {
            var text = txtOutput.Text ?? string.Empty;
            const string delimiter = DELIMITER;
            var index = text.IndexOf( delimiter, StringComparison.Ordinal );
            if ( index >= 0 )
            {
                var start = index + delimiter.Length;
                while ( start < text.Length && ( text[start] == '\r' || text[start] == '\n' ) )
                {
                    start++;
                }

                text = text.Substring( start );
            }

            var lengthText = ( text?.Length ?? 0 ).ToString( "#,0" );
            lblLength.Text = $"文字数: {lengthText}";
        }

        private void SetBuildRunning ( bool running )
        {
            btnRun.Enabled = !running;
            btnCancel.Enabled = running;
            btnRefDirPath.Enabled = !running;
            chkIgnoreExtNegated.Enabled = !running;
            chkOutputToFile.Enabled = !running;
            progressBar1.Style = ProgressBarStyle.Blocks;
            if ( !running )
            {
                progressBar1.Value = 0;
            }
        }

        private void UpdateBuildProgress ( BuildProgress progress )
        {
            if ( progress.IsIndeterminate )
            {
                progressBar1.Style = ProgressBarStyle.Marquee;
                lblProgress.Text = progress.Message;
                return;
            }

            progressBar1.Style = ProgressBarStyle.Blocks;
            progressBar1.Maximum = Math.Max( 1, progress.Total );
            progressBar1.Value = Math.Min( progress.Current, progressBar1.Maximum );
            lblProgress.Text = progress.Message;
        }

        private class FileReadResult
        {
            public string FilePath { get; set; }
            public string Content { get; set; }
            public string EncodingName { get; set; }
            public string LineEnding { get; set; }
        }

        private class BuildResult
        {
            public List<FileReadResult> Files { get; set; }
            public string OutputPath { get; set; }
            public int TargetCount { get; set; }
            public int OutputCount { get; set; }
            public int SkippedCount { get; set; }
        }

        private class BuildProgress
        {
            public bool IsIndeterminate { get; private set; }
            public int Current { get; private set; }
            public int Total { get; private set; }
            public string Message { get; private set; }

            public static BuildProgress Indeterminate ( string message )
            {
                return new BuildProgress { IsIndeterminate = true, Message = message };
            }

            public static BuildProgress Determinate ( int current, int total, string message )
            {
                return new BuildProgress { Current = current, Total = total, Message = message };
            }
        }
    }
}
