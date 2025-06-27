using System;
using System.Drawing;
using System.Windows.Forms;

namespace CursorMover
{
    public partial class MainForm : Form
    {
        // 次にマウスカーソルを右に動かすかどうかのフラグ
        private bool moveRight = true;

        public MainForm()
        {
            InitializeComponent();

            // テキストボックスの初期値を設定
            this.textCursorMoveTimeSec.Text = "5";

            // イベントハンドラを登録
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
        }

        /// <summary>
        /// STARTボタンクリック時の処理
        /// </summary>
        private void btnStart_Click(object sender, EventArgs e)
        {
            int intervalSeconds;
            // テキストボックスの値が正の整数であるか検証
            if (int.TryParse(textCursorMoveTimeSec.Text, out intervalSeconds) && intervalSeconds > 0)
            {
                // タイマーの間隔をミリ秒で設定
                timer1.Interval = intervalSeconds * 1000;
                // タイマーを開始
                timer1.Start();

                // UIの状態を更新
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                textCursorMoveTimeSec.ReadOnly = true;
            }
            else
            {
                // 無効な値の場合はメッセージを表示
                MessageBox.Show("間隔には正の整数を秒単位で入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// STOPボタンクリック時の処理
        /// </summary>
        private void btnStop_Click(object sender, EventArgs e)
        {
            // タイマーを停止
            timer1.Stop();

            // UIの状態を更新
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            textCursorMoveTimeSec.ReadOnly = false;
        }

        /// <summary>
        /// タイマーイベント発生時の処理
        /// </summary>
        private void timer1_Tick(object sender, EventArgs e)
        {
            // 現在のマウスカーソル位置を取得
            Point currentPosition = Cursor.Position;

            if (moveRight)
            {
                // 5ピクセル右に移動
                Cursor.Position = new Point(currentPosition.X + 5, currentPosition.Y);
            }
            else
            {
                // 5ピクセル左に移動
                Cursor.Position = new Point(currentPosition.X - 5, currentPosition.Y);
            }

            // 次の移動方向を反転させる
            moveRight = !moveRight;
        }
    }
}