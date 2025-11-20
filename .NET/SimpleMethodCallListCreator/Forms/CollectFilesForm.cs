using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace SimpleMethodCallListCreator.Forms
{
    public partial class CollectFilesForm : Form
    {
        private readonly AppSettings _settings;
        private readonly CollectFilesService _service = new CollectFilesService();
        private CancellationTokenSource _cancellationTokenSource;
        private bool _isRunning;
        private bool _isProgressDeterminate;

        public CollectFilesForm ( AppSettings settings )
        {
            _settings = settings ?? new AppSettings();
            InitializeComponent();
            InitializeControls();
            HookEvents();
            LoadSettings();
        }

        private void InitializeControls ()
        {
            lblStatus.Text = string.Empty;
            UpdateFailedLabel( 0 );
            pbProgress.Visible = false;
            pbProgress.Style = ProgressBarStyle.Blocks;
            pbProgress.Minimum = 0;
            pbProgress.Maximum = 1;
            pbProgress.Value = 0;
            btnCancel.Enabled = false;
            txtMethodListPath.AllowDrop = true;
            txtStartSrcFilePath.AllowDrop = true;
            txtCollectDirPath.AllowDrop = true;
        }

        private void HookEvents ()
        {
            btnRun.Click += BtnRun_Click;
            btnCancel.Click += BtnCancel_Click;
            btnRefMethodListPath.Click += BtnRefMethodListPath_Click;
            btnRefStartSrcFilePath.Click += BtnRefStartSrcFilePath_Click;
            btnRefCollectDirPath.Click += BtnRefCollectDirPath_Click;
            txtMethodListPath.DragEnter += FilePathTextBox_DragEnter;
            txtStartSrcFilePath.DragEnter += FilePathTextBox_DragEnter;
            txtCollectDirPath.DragEnter += FilePathTextBox_DragEnter;
            txtMethodListPath.DragDrop += TxtMethodListPath_DragDrop;
            txtStartSrcFilePath.DragDrop += TxtStartSrcFilePath_DragDrop;
            txtCollectDirPath.DragDrop += TxtCollectDirPath_DragDrop;
            FormClosing += CollectFilesForm_FormClosing;
        }

        private void LoadSettings ()
        {
            txtMethodListPath.Text = _settings.LastCollectMethodListPath ?? string.Empty;
            txtStartSrcFilePath.Text = _settings.LastCollectSourceFilePath ?? string.Empty;
            txtStartMethod.Text = _settings.LastCollectMethod ?? string.Empty;
            txtCollectDirPath.Text = _settings.LastCollectTargetDirectory ?? string.Empty;
        }

        private void SaveSettings ()
        {
            _settings.LastCollectMethodListPath = ( txtMethodListPath.Text ?? string.Empty ).Trim();
            _settings.LastCollectSourceFilePath = ( txtStartSrcFilePath.Text ?? string.Empty ).Trim();
            _settings.LastCollectMethod = ( txtStartMethod.Text ?? string.Empty ).Trim();
            _settings.LastCollectTargetDirectory = ( txtCollectDirPath.Text ?? string.Empty ).Trim();
            SettingsManager.Save( _settings );
        }

        private void CollectFilesForm_FormClosing ( object sender, FormClosingEventArgs e )
        {
            if ( _isRunning )
            {
                _cancellationTokenSource?.Cancel();
            }

            SaveSettings();
        }

        private async void BtnRun_Click ( object sender, EventArgs e )
        {
            if ( _isRunning )
            {
                return;
            }

            await ExecuteCollectAsync();
        }

        private void BtnCancel_Click ( object sender, EventArgs e )
        {
            if ( !_isRunning )
            {
                return;
            }

            btnCancel.Enabled = false;
            _cancellationTokenSource?.Cancel();
        }

        private async Task ExecuteCollectAsync ()
        {
            var methodListPath = ( txtMethodListPath.Text ?? string.Empty ).Trim();
            var startSourceFilePath = ( txtStartSrcFilePath.Text ?? string.Empty ).Trim();
            var startMethod = ( txtStartMethod.Text ?? string.Empty ).Trim();
            var collectDirPath = ( txtCollectDirPath.Text ?? string.Empty ).Trim();

            if ( methodListPath.Length == 0 )
            {
                MessageBox.Show( this, "メソッドリストのパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( !File.Exists( methodListPath ) )
            {
                MessageBox.Show( this, "指定されたメソッドリストが見つかりません。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( !Path.IsPathRooted( methodListPath ) )
            {
                MessageBox.Show( this, "メソッドリストのパスは絶対パスで指定してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( startSourceFilePath.Length == 0 )
            {
                MessageBox.Show( this, "開始ソースファイルのパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( !File.Exists( startSourceFilePath ) )
            {
                MessageBox.Show( this, "開始ソースファイルが見つかりません。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( !string.Equals( Path.GetExtension( startSourceFilePath ), ".java", StringComparison.OrdinalIgnoreCase ) )
            {
                MessageBox.Show( this, "Javaファイル（*.java）のみが対象です。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( startMethod.Length == 0 )
            {
                MessageBox.Show( this, "開始メソッドを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            if ( collectDirPath.Length == 0 )
            {
                MessageBox.Show( this, "収集フォルダパスを入力してください。", "入力エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            BeginExecution();
            var progress = new Progress<CollectFilesService.CollectFilesProgressInfo>( UpdateProgress );
            var token = _cancellationTokenSource.Token;

            try
            {
                var result = await Task.Run( () =>
                    _service.CollectFiles( methodListPath, startSourceFilePath, startMethod, collectDirPath, progress,
                        token ), token );

                SaveSettings();

                var message = new StringBuilder();
                message.AppendLine( "処理が完了しました。" );
                message.AppendLine( $"対象ファイル数: {result.TotalCount}" );
                message.AppendLine( $"コピー済: {result.CopiedCount}" );
                message.AppendLine( $"スキップ: {result.SkippedCount}" );
                if ( result.FailureCount > 0 )
                {
                    message.AppendLine( $"メソッド特定失敗: {result.FailureCount}" );
                }
                MessageBox.Show( this, message.ToString(), "完了", MessageBoxButtons.OK, MessageBoxIcon.Information );
                if ( result.FailureCount > 0 )
                {
                    LogFailureDetails( methodListPath, startSourceFilePath, startMethod, collectDirPath, result );
                }
                UpdateFailedLabel( result.FailureCount );
            }
            catch ( OperationCanceledException )
            {
                MessageBox.Show( this, "処理を中止しました。", "中止",
                    MessageBoxButtons.OK, MessageBoxIcon.Information );
            }
            catch ( JavaParseException ex )
            {
                var builder = new StringBuilder();
                builder.AppendLine( "Javaファイルの解析に失敗しました。" );
                builder.AppendLine( $"行番号: {ex.LineNumber}" );
                if ( !string.IsNullOrEmpty( ex.InvalidContent ) )
                {
                    builder.AppendLine( $"内容: {ex.InvalidContent}" );
                }

                MessageBox.Show( this, builder.ToString(), "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
                LogErrorDetail( methodListPath, startSourceFilePath, startMethod, collectDirPath, ex );
                UpdateFailedLabel( 1 );
            }
            catch ( CollectFilesService.MethodAmbiguityException ex )
            {
                MessageBox.Show( this, ex.Message, "メソッド特定エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
                LogErrorDetail( methodListPath, startSourceFilePath, startMethod, collectDirPath, ex );
                UpdateFailedLabel( 1 );
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"ファイル収集に失敗しました。\n{ex.Message}", "エラー",
                    MessageBoxButtons.OK, MessageBoxIcon.Error );
                LogErrorDetail( methodListPath, startSourceFilePath, startMethod, collectDirPath, ex );
                UpdateFailedLabel( 1 );
            }
            finally
            {
                EndExecution();
            }
        }

        private void BeginExecution ()
        {
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = new CancellationTokenSource();
            _isRunning = true;
            _isProgressDeterminate = false;
            lblStatus.Text = string.Empty;
            UpdateFailedLabel( 0 );
            SetRunningState( true, ProgressBarStyle.Marquee );
        }

        private void EndExecution ()
        {
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = null;
            _isRunning = false;
            _isProgressDeterminate = false;
            SetRunningState( false, ProgressBarStyle.Blocks );
        }

        private void SetRunningState ( bool isRunning, ProgressBarStyle style )
        {
            btnRun.Enabled = !isRunning;
            btnCancel.Enabled = isRunning;
            txtMethodListPath.Enabled = !isRunning;
            btnRefMethodListPath.Enabled = !isRunning;
            txtStartSrcFilePath.Enabled = !isRunning;
            btnRefStartSrcFilePath.Enabled = !isRunning;
            txtStartMethod.Enabled = !isRunning;
            txtCollectDirPath.Enabled = !isRunning;
            btnRefCollectDirPath.Enabled = !isRunning;

            pbProgress.Visible = isRunning;
            pbProgress.Style = style;
            pbProgress.Value = pbProgress.Minimum;

            Cursor = isRunning ? Cursors.WaitCursor : Cursors.Default;
        }

        private void UpdateProgress ( CollectFilesService.CollectFilesProgressInfo info )
        {
            if ( info == null )
            {
                return;
            }

            if ( !_isProgressDeterminate )
            {
                pbProgress.Style = ProgressBarStyle.Continuous;
                pbProgress.Minimum = 0;
                pbProgress.Maximum = Math.Max( 1, info.TotalFileCount );
                pbProgress.Value = 0;
                _isProgressDeterminate = true;
            }

            var processed = info.CopiedCount + info.SkippedCount;
            pbProgress.Value = Math.Min( Math.Max( processed, pbProgress.Minimum ), pbProgress.Maximum );
            lblStatus.Text = $"{info.CopiedCount + info.SkippedCount:D4}/{info.TotalFileCount:D4}ファイル";
        }

        private void BtnRefMethodListPath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new OpenFileDialog() )
            {
                dialog.Filter = "TSV ファイル (*.tsv)|*.tsv|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                var currentPath = ( txtMethodListPath.Text ?? string.Empty ).Trim();
                if ( currentPath.Length > 0 )
                {
                    var directory = Path.GetDirectoryName( currentPath );
                    if ( !string.IsNullOrEmpty( directory ) && Directory.Exists( directory ) )
                    {
                        dialog.InitialDirectory = directory;
                    }
                }

                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtMethodListPath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRefStartSrcFilePath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new OpenFileDialog() )
            {
                dialog.Filter = "Java ファイル (*.java)|*.java|すべてのファイル (*.*)|*.*";
                dialog.Multiselect = false;
                var currentPath = ( txtStartSrcFilePath.Text ?? string.Empty ).Trim();
                if ( currentPath.Length > 0 )
                {
                    var directory = Path.GetDirectoryName( currentPath );
                    if ( !string.IsNullOrEmpty( directory ) && Directory.Exists( directory ) )
                    {
                        dialog.InitialDirectory = directory;
                    }
                }

                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtStartSrcFilePath.Text = dialog.FileName;
                }
            }
        }

        private void BtnRefCollectDirPath_Click ( object sender, EventArgs e )
        {
            using ( var dialog = new FolderBrowserDialog() )
            {
                dialog.Description = "収集フォルダを選択してください。";
                var currentPath = ( txtCollectDirPath.Text ?? string.Empty ).Trim();
                if ( currentPath.Length > 0 && Directory.Exists( currentPath ) )
                {
                    dialog.SelectedPath = currentPath;
                }

                if ( dialog.ShowDialog( this ) == DialogResult.OK )
                {
                    txtCollectDirPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void FilePathTextBox_DragEnter ( object sender, DragEventArgs e )
        {
            if ( e.Data.GetDataPresent( DataFormats.FileDrop ) )
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TxtMethodListPath_DragDrop ( object sender, DragEventArgs e )
        {
            var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
            if ( files == null || files.Length == 0 )
            {
                return;
            }

            txtMethodListPath.Text = files[0];
        }

        private void TxtStartSrcFilePath_DragDrop ( object sender, DragEventArgs e )
        {
            var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
            if ( files == null || files.Length == 0 )
            {
                return;
            }

            txtStartSrcFilePath.Text = files[0];
        }

        private void TxtCollectDirPath_DragDrop ( object sender, DragEventArgs e )
        {
            var files = e.Data.GetData( DataFormats.FileDrop ) as string[];
            if ( files == null || files.Length == 0 )
            {
                return;
            }

            var path = files[0];
            if ( Directory.Exists( path ) )
            {
                txtCollectDirPath.Text = path;
                return;
            }

            var directory = Path.GetDirectoryName( path );
            if ( !string.IsNullOrEmpty( directory ) && Directory.Exists( directory ) )
            {
                txtCollectDirPath.Text = directory;
            }
        }

        private static void LogErrorDetail ( string methodListPath, string startSourceFilePath, string startMethod,
            string collectDirPath, Exception ex )
        {
            var builder = new StringBuilder();
            builder.AppendLine( "ファイル収集中にエラーが発生しました。" );
            builder.AppendLine( $"メソッドリスト: {methodListPath}" );
            builder.AppendLine( $"開始ソースファイル: {startSourceFilePath}" );
            builder.AppendLine( $"開始メソッド: {startMethod}" );
            builder.AppendLine( $"収集フォルダ: {collectDirPath}" );
            if ( ex is JavaParseException parseEx )
            {
                builder.AppendLine( $"行番号: {parseEx.LineNumber}" );
                if ( !string.IsNullOrEmpty( parseEx.InvalidContent ) )
                {
                    builder.AppendLine( $"内容: {parseEx.InvalidContent}" );
                }
            }

            if ( ex is CollectFilesService.MethodAmbiguityException ambiguityEx &&
                 ambiguityEx.Candidates != null &&
                 ambiguityEx.Candidates.Count > 0 )
            {
                builder.AppendLine( "候補一覧:" );
                foreach ( var candidate in ambiguityEx.Candidates )
                {
                    builder.AppendLine( $" - {candidate}" );
                }
            }

            builder.AppendLine( "例外詳細:" );
            builder.AppendLine( ex.ToString() );
            ErrorLogger.LogError( builder.ToString() );
        }

        private static void LogFailureDetails ( string methodListPath, string startSourceFilePath, string startMethod,
            string collectDirPath, CollectFilesService.CollectFilesResult result )
        {
            if ( result == null || result.FailureCount <= 0 )
            {
                return;
            }

            var builder = new StringBuilder();
            builder.AppendLine( "呼び出し先メソッドを特定できない箇所がありました。" );
            builder.AppendLine( $"メソッドリスト: {methodListPath}" );
            builder.AppendLine( $"開始ソースファイル: {startSourceFilePath}" );
            builder.AppendLine( $"開始メソッド: {startMethod}" );
            builder.AppendLine( $"収集フォルダ: {collectDirPath}" );
            builder.AppendLine( $"失敗件数: {result.FailureCount}" );
            builder.AppendLine( "詳細:" );

            foreach ( var detail in result.FailureDetails )
            {
                if ( string.IsNullOrWhiteSpace( detail ) )
                {
                    continue;
                }

                builder.AppendLine( detail );
                builder.AppendLine();
            }

            ErrorLogger.LogError( builder.ToString().TrimEnd() );
        }

        private void UpdateFailedLabel ( int failureCount )
        {
            if ( failureCount > 0 )
            {
                lblFailed.Text = $"メソッド特定失敗：{failureCount}件";
                lblFailed.Visible = true;
            }
            else
            {
                lblFailed.Visible = false;
            }
        }
    }
}
