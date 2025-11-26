using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// アプリケーションのログ記録を担当するサービス
    /// </summary>
    public class LogService
    {
        private readonly string _logFilePath;
        public bool IsLoggingEnabled { get; set; }
        private readonly Label _statusLabel;

        /// <summary>
        /// LogServiceのコンストラクタ
        /// </summary>
        /// <param name="statusLabel">ステータス表示に使用するUIラベル</param>
        /// <param name="isLoggingEnabled">ログ記録を有効にするかどうか</param>
        public LogService(Label statusLabel, bool isLoggingEnabled = true)
        {
            _statusLabel = statusLabel;
            IsLoggingEnabled = isLoggingEnabled;
            _logFilePath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "SimpleExcelGrep_log.txt");
        }

        /// <summary>
        /// メッセージをログに記録し、オプションでステータスラベルに表示
        /// </summary>
        public void LogMessage(string message, bool showInStatus = false, bool force = false)
        {
            if (!IsLoggingEnabled && !force) return;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string logEntry = $"[{timestamp}] {message}";

            try
            {
                // ログファイルに追記
                File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);

                // デバッグコンソールにも出力
                Debug.WriteLine(logEntry);

                // オプションでステータスバーに表示
                if (showInStatus)
                {
                    UpdateStatus(message);
                }
            }
            catch (Exception ex)
            {
                // ログ出力自体が失敗した場合、ステータスバーだけに表示
                UpdateStatus($"ログの記録に失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 環境情報をログに記録
        /// </summary>
        public void LogEnvironmentInfo()
        {
            LogMessage("======== 環境情報 ========");
            LogMessage($"OSバージョン: {Environment.OSVersion}");
            LogMessage($".NETバージョン: {Environment.Version}");
            LogMessage($"64ビットOS: {Environment.Is64BitOperatingSystem}");
            LogMessage($"64ビットプロセス: {Environment.Is64BitProcess}");
            LogMessage($"マシン名: {Environment.MachineName}");

            // Officeのバージョン情報取得を試行（Late Binding方式）
            try
            {
                // Excelアプリケーションを作成
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType != null)
                {
                    object excelApp = Activator.CreateInstance(excelType);

                    // バージョン情報を取得
                    object version = excelType.InvokeMember("Version",
                        System.Reflection.BindingFlags.GetProperty,
                        null, excelApp, null);
                    LogMessage($"Excel バージョン: {version}");

                    // ビルド情報を取得
                    object build = excelType.InvokeMember("Build",
                        System.Reflection.BindingFlags.GetProperty,
                        null, excelApp, null);
                    LogMessage($"Excel ビルド: {build}");

                    // COMの早期解放
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); } catch { }
                }
                else
                {
                    LogMessage("Excel.ApplicationのCOMタイプが見つかりません");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Excelバージョン取得エラー: {ex.Message}");
            }

            LogMessage("==========================");
        }

        /// <summary>
        /// ステータスラベルのテキストを更新
        /// </summary>
        public void UpdateStatus(string message)
        {
            if (_statusLabel == null) return;

            if (_statusLabel.InvokeRequired)
            {
                _statusLabel.Invoke(new Action(() => _statusLabel.Text = message));
            }
            else
            {
                _statusLabel.Text = message;
            }
        }
    }
}
