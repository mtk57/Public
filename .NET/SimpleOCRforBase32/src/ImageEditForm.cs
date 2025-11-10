using System;
using System.IO;
using System.Windows.Forms;
using OpenCvSharp;
using OpenCvSharp.Extensions;

namespace SimpleOCRforBase32
{
    public partial class ImageEditForm : Form
    {
        private const double DefaultContrast = 1.0;
        private const int DefaultBrightness = 0;
        private const int DefaultBinaryThreshold = 128;

        private readonly string imagePath;
        private bool suppressEvents;
        private Mat originalMat;
        private Mat currentProcessedMat;

        public ImageEditForm ( string imagePath )
        {
            this.imagePath = imagePath ?? throw new ArgumentNullException( nameof( imagePath ) );
            InitializeComponent();
            InitializeControls();
        }

        public string EditedImagePath { get; private set; }

        protected override void OnLoad ( EventArgs e )
        {
            base.OnLoad( e );

            if ( !LoadOriginalImage() )
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }

            UpdatePreview();
        }

        protected override void OnFormClosed ( EventArgs e )
        {
            base.OnFormClosed( e );
            DisposeResources();
        }

        private void InitializeControls ()
        {
            picPreview.SizeMode = PictureBoxSizeMode.Zoom;

            trackBarContrast.Minimum = 50;
            trackBarContrast.Maximum = 200;
            trackBarContrast.TickFrequency = 10;
            trackBarContrast.Value = (int)( DefaultContrast * 100 );

            trackBarBrightness.Minimum = -100;
            trackBarBrightness.Maximum = 100;
            trackBarBrightness.TickFrequency = 10;
            trackBarBrightness.Value = DefaultBrightness;

            trackBarBinary.Minimum = 0;
            trackBarBinary.Maximum = 255;
            trackBarBinary.TickFrequency = 5;
            trackBarBinary.Value = DefaultBinaryThreshold;

            trackBarContrast.Scroll += TrackBar_ValueChanged;
            trackBarBrightness.Scroll += TrackBar_ValueChanged;
            trackBarBinary.Scroll += TrackBar_ValueChanged;

            btnDefault.Click += BtnDefault_Click;
            btnSave.Click += BtnSave_Click;
            btnCancel.Click += BtnCancel_Click;
        }

        private bool LoadOriginalImage ()
        {
            try
            {
                originalMat?.Dispose();
                originalMat = Cv2.ImRead( imagePath, ImreadModes.Color );
                if ( originalMat.Empty() )
                {
                    throw new InvalidOperationException( "画像を読み込めませんでした。" );
                }

                return true;
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"画像の読み込みに失敗しました。\n{ex.Message}", "読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
                originalMat?.Dispose();
                originalMat = null;
                return false;
            }
        }

        private void TrackBar_ValueChanged ( object sender, EventArgs e )
        {
            if ( suppressEvents )
            {
                return;
            }

            UpdatePreview();
        }

        private void UpdatePreview ()
        {
            if ( originalMat == null || originalMat.Empty() )
            {
                return;
            }

            try
            {
                using ( var gray = new Mat() )
                using ( var adjusted = new Mat() )
                using ( var binary = new Mat() )
                using ( var display = new Mat() )
                {
                    Cv2.CvtColor( originalMat, gray, ColorConversionCodes.BGR2GRAY );
                    var contrast = Math.Max( 0.1, trackBarContrast.Value / 100.0 );
                    var brightness = trackBarBrightness.Value;
                    gray.ConvertTo( adjusted, MatType.CV_8U, contrast, brightness );
                    var threshold = trackBarBinary.Value;
                    Cv2.Threshold( adjusted, binary, threshold, 255, ThresholdTypes.Binary );
                    Cv2.CvtColor( binary, display, ColorConversionCodes.GRAY2BGR );

                    currentProcessedMat?.Dispose();
                    currentProcessedMat = binary.Clone();

                    var previewBitmap = BitmapConverter.ToBitmap( display );
                    var previousImage = picPreview.Image;
                    picPreview.Image = previewBitmap;
                    previousImage?.Dispose();
                }
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"画像のプレビュー更新に失敗しました。\n{ex.Message}", "処理エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        private void BtnDefault_Click ( object sender, EventArgs e )
        {
            var result = MessageBox.Show( this, "スライダーを初期値に戻しますか？", "デフォルトに戻す", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
            if ( result != DialogResult.Yes )
            {
                return;
            }

            suppressEvents = true;
            trackBarContrast.Value = (int)( DefaultContrast * 100 );
            trackBarBrightness.Value = DefaultBrightness;
            trackBarBinary.Value = DefaultBinaryThreshold;
            suppressEvents = false;
            UpdatePreview();
        }

        private void BtnSave_Click ( object sender, EventArgs e )
        {
            if ( currentProcessedMat == null || currentProcessedMat.Empty() )
            {
                MessageBox.Show( this, "処理後の画像がありません。", "保存不可", MessageBoxButtons.OK, MessageBoxIcon.Warning );
                return;
            }

            var result = MessageBox.Show( this, "調整後の画像を保存してメイン画面に戻りますか？", "保存確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
            if ( result != DialogResult.Yes )
            {
                return;
            }

            try
            {
                var outputPath = GenerateEditedFilePath();
                Cv2.ImWrite( outputPath, currentProcessedMat );
                EditedImagePath = outputPath;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch ( Exception ex )
            {
                MessageBox.Show( this, $"画像の保存に失敗しました。\n{ex.Message}", "保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Error );
            }
        }

        private void BtnCancel_Click ( object sender, EventArgs e )
        {
            var result = MessageBox.Show( this, "編集内容を破棄してメイン画面に戻りますか？", "キャンセル確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question );
            if ( result == DialogResult.Yes )
            {
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        private string GenerateEditedFilePath ()
        {
            var directory = Path.GetDirectoryName( imagePath );
            if ( string.IsNullOrEmpty( directory ) )
            {
                directory = Environment.CurrentDirectory;
            }

            var baseName = Path.GetFileNameWithoutExtension( imagePath );
            var extension = Path.GetExtension( imagePath );

            for ( var i = 1; i <= 999; i++ )
            {
                var candidate = Path.Combine( directory, $"{baseName}_edit{i:000}{extension}" );
                if ( !File.Exists( candidate ) )
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException( "保存可能なファイル名が見つかりませんでした。（_edit001～_edit999）" );
        }

        private void DisposeResources ()
        {
            picPreview.Image?.Dispose();
            picPreview.Image = null;
            currentProcessedMat?.Dispose();
            currentProcessedMat = null;
            originalMat?.Dispose();
            originalMat = null;
        }
    }
}
