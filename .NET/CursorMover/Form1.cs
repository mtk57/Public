using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices; // DllImportに必要
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace CursorMover
{
    public partial class MainForm : Form
    {
        // 次にマウスカーソルを右に動かすかどうかのフラグ
        private bool moveRight = true;
        private readonly Stopwatch elapsedStopwatch = new Stopwatch();
        private string elapsedTimeBeforeTextBeforeEdit = string.Empty;
        private readonly string settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CursorMoverSettings.json");
        private static readonly Regex ElapsedTimePattern = new Regex(@"^\d{2,}:[0-5]\d:[0-5]\d$", RegexOptions.Compiled);

        public MainForm()
        {
            InitializeComponent();

            // テキストボックスの初期値を設定
            this.textCursorMoveTimeSec.Text = "5";

            // イベントハンドラを登録
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            this.btnStop.Click += new System.EventHandler(this.btnStop_Click);
            this.btnSaveElapsedTimeBefore.Click += new System.EventHandler(this.btnSaveElapsedTimeBefore_Click);
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            this.txtElapsedTimeBefore.Enter += new System.EventHandler(this.txtElapsedTimeBefore_Enter);
            this.txtElapsedTimeBefore.Leave += new System.EventHandler(this.txtElapsedTimeBefore_Leave);
            this.FormClosing += new FormClosingEventHandler(this.MainForm_FormClosing);
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
                elapsedStopwatch.Restart();
                lblElapsedTime.Text = "00:00:00";

                // UIの状態を更新
                btnStart.Enabled = false;
                btnStop.Enabled = true;
                btnSaveElapsedTimeBefore.Enabled = false;
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
            elapsedStopwatch.Stop();
            lblElapsedTime.Text = FormatElapsedTime(elapsedStopwatch.Elapsed);

            // UIの状態を更新
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            btnSaveElapsedTimeBefore.Enabled = true;
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

            // 画面初期状態を設定
            btnSaveElapsedTimeBefore.Enabled = false;
            lblElapsedTime.Text = "00:00:00";

            // 前回終了時のUI状態を復元
            LoadUiStateFromJson();
            elapsedTimeBeforeTextBeforeEdit = txtElapsedTimeBefore.Text;
        }

        private void btnSaveElapsedTimeBefore_Click(object sender, EventArgs e)
        {
            txtElapsedTimeBefore.Text = lblElapsedTime.Text;
            elapsedTimeBeforeTextBeforeEdit = txtElapsedTimeBefore.Text;
        }

        private void txtElapsedTimeBefore_Enter(object sender, EventArgs e)
        {
            elapsedTimeBeforeTextBeforeEdit = txtElapsedTimeBefore.Text;
        }

        private void txtElapsedTimeBefore_Leave(object sender, EventArgs e)
        {
            string currentText = txtElapsedTimeBefore.Text.Trim();

            // 値が変わっていない場合は何もしない
            if (string.Equals(currentText, elapsedTimeBeforeTextBeforeEdit, StringComparison.Ordinal))
            {
                return;
            }

            if (!IsValidElapsedTimeText(currentText))
            {
                txtElapsedTimeBefore.Text = elapsedTimeBeforeTextBeforeEdit;
                MessageBox.Show("経過時間（前回）は HH:mm:ss 形式で入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            txtElapsedTimeBefore.Text = currentText;
            elapsedTimeBeforeTextBeforeEdit = currentText;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveUiStateToJson();
        }

        private static string FormatElapsedTime(TimeSpan elapsed)
        {
            int totalHours = (int)elapsed.TotalHours;
            return string.Format(CultureInfo.InvariantCulture, "{0:D2}:{1:D2}:{2:D2}", totalHours, elapsed.Minutes, elapsed.Seconds);
        }

        private static bool IsValidElapsedTimeText(string value)
        {
            return ElapsedTimePattern.IsMatch(value ?? string.Empty);
        }

        private void LoadUiStateFromJson()
        {
            if (!File.Exists(settingsFilePath))
            {
                return;
            }

            try
            {
                string json = File.ReadAllText(settingsFilePath, Encoding.UTF8);
                var serializer = new JavaScriptSerializer();
                UiState uiState = serializer.Deserialize<UiState>(json);

                if (uiState == null)
                {
                    return;
                }

                if (!string.IsNullOrWhiteSpace(uiState.CursorMoveIntervalSeconds))
                {
                    textCursorMoveTimeSec.Text = uiState.CursorMoveIntervalSeconds.Trim();
                }

                if (!string.IsNullOrWhiteSpace(uiState.ElapsedTimeBefore))
                {
                    string elapsedText = uiState.ElapsedTimeBefore.Trim();
                    if (IsValidElapsedTimeText(elapsedText))
                    {
                        txtElapsedTimeBefore.Text = elapsedText;
                    }
                }

                if (uiState.Memo != null)
                {
                    txtMemo.Text = uiState.Memo;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの読み込みに失敗しました。\n{ex.Message}", "設定読込エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SaveUiStateToJson()
        {
            try
            {
                string elapsedTimeBefore = txtElapsedTimeBefore.Text?.Trim() ?? string.Empty;
                if (!IsValidElapsedTimeText(elapsedTimeBefore))
                {
                    elapsedTimeBefore = elapsedTimeBeforeTextBeforeEdit;
                }

                var uiState = new UiState
                {
                    CursorMoveIntervalSeconds = textCursorMoveTimeSec.Text?.Trim(),
                    ElapsedTimeBefore = elapsedTimeBefore,
                    Memo = txtMemo.Text ?? string.Empty
                };

                var serializer = new JavaScriptSerializer();
                string json = serializer.Serialize(uiState);
                File.WriteAllText(settingsFilePath, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルの保存に失敗しました。\n{ex.Message}", "設定保存エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private sealed class UiState
        {
            public string CursorMoveIntervalSeconds { get; set; }
            public string ElapsedTimeBefore { get; set; }
            public string Memo { get; set; }
        }

    }
}
