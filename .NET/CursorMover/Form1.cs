using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices; // DllImportに必要
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
            // 移動量を決定
            int dx = moveRight ? 5 : -5;
            
            // SendInput APIを使ってマウスの移動をシミュレート
            SimulateMouseMovement(dx, 0);

            // 次の移動方向を反転させる
            moveRight = !moveRight;
        }

        /// <summary>
        /// マウスの移動をシミュレートします。
        /// </summary>
        /// <param name="dx">X方向の相対移動量</param>
        /// <param name="dy">Y方向の相対移動量</param>
        private void SimulateMouseMovement(int dx, int dy)
        {
            var input = new INPUT
            {
                type = INPUT_MOUSE,
                u = new InputUnion
                {
                    mi = new MOUSEINPUT
                    {
                        dx = dx,
                        dy = dy,
                        mouseData = 0,
                        dwFlags = MOUSEEVENTF_MOVE,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            // INPUT構造体の配列を作成
            INPUT[] inputs = { input };

            // SendInputを呼び出してマウスイベントを送信
            SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(INPUT)));
        }

        // --- Win32 API 定義 ---

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        private const int INPUT_MOUSE = 0;
        private const int MOUSEEVENTF_MOVE = 0x0001;

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public int type;
            public InputUnion u;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private void MainForm_Load ( object sender, EventArgs e )
        {
            // タイトルにバージョン情報を表示
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"{this.Text}  ver {version.Major}.{version.Minor}.{version.Build}";
        }
    }
}