using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text; // For StringBuilder
using System.Threading; // For Thread.CurrentThread
using System.Windows.Forms;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// アプリケーションのログ記録を担当するサービス
    /// </summary>
    internal class LogService
    {
        private readonly string _logFilePath;
        private string _performanceLogFilePath; // TSVログファイルパス
        private readonly bool _isLoggingEnabled;
        private readonly Label _statusLabel;
        private readonly object _logLock = new object(); // 通常ログファイルアクセス用ロックオブジェクト
        private readonly object _performanceLogLock = new object(); // TSVログファイルアクセス用ロックオブジェクト

        /// <summary>
        /// LogServiceのコンストラクタ
        /// </summary>
        /// <param name="statusLabel">ステータス表示に使用するUIラベル</param>
        /// <param name="isLoggingEnabled">ログ記録を有効にするかどうか</param>
        public LogService(Label statusLabel, bool isLoggingEnabled = true)
        {
            _statusLabel = statusLabel;
            _isLoggingEnabled = isLoggingEnabled;
            string logDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            _logFilePath = Path.Combine(logDirectory, "SimpleExcelGrep_log.txt");

            if (_isLoggingEnabled)
            {
                // アプリケーション起動時に一度だけログファイルの内容をクリア（または追記のままにするか選択）
                // File.WriteAllText(_logFilePath, string.Empty, Encoding.UTF8); // クリアする場合
            }
        }

        /// <summary>
        /// パフォーマンスログファイル名を初期化し、ヘッダーを書き込む
        /// </summary>
        public void InitializePerformanceLog()
        {
            if (!_isLoggingEnabled) return;

            string logDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            _performanceLogFilePath = Path.Combine(logDirectory, $"SimpleExcelGrep_Performance_{timestamp}.tsv");

            try
            {
                // ヘッダー行を書き込む
                string header = "Timestamp\tProcessedFileCount\tElapsedTimeSeconds\tMemoryUsageMB\tFilePath\tFileSizeMB\tProcessingTimeMs";
                lock (_performanceLogLock)
                {
                    File.WriteAllText(_performanceLogFilePath, header + Environment.NewLine, Encoding.UTF8);
                }
                LogMessage($"パフォーマンスログを初期化しました: {_performanceLogFilePath}");
            }
            catch (Exception ex)
            {
                LogMessage($"パフォーマンスログの初期化に失敗: {ex.Message}", true);
            }
        }

        /// <summary>
        /// パフォーマンスデータをTSVファイルに記録
        /// </summary>
        /// <param name="processedFileCount">処理済みファイル数</param>
        /// <param name="elapsedTimeSeconds">検索開始からの経過秒数</param>
        /// <param name="memoryUsageMB">現在のメモリ使用量 (MB)</param>
        /// <param name="filePathProcessed">処理したファイルパス</param>
        /// <param name="fileSizeMB">処理したファイルのサイズ (MB)</param>
        /// <param name="fileProcessingTimeMs">そのファイルの処理にかかった時間 (ミリ秒)</param>
        public void LogPerformanceData(int processedFileCount, double elapsedTimeSeconds, double memoryUsageMB, string filePathProcessed, double fileSizeMB, double fileProcessingTimeMs)
        {
            if (!_isLoggingEnabled || string.IsNullOrEmpty(_performanceLogFilePath)) return;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            // TSVなので、タブや改行が含まれる可能性のある filePathProcessed はエスケープした方が安全
            string escapedFilePath = EscapeTsvField(filePathProcessed);
            string logEntry = $"{timestamp}\t{processedFileCount}\t{elapsedTimeSeconds:F3}\t{memoryUsageMB:F2}\t{escapedFilePath}\t{fileSizeMB:F2}\t{fileProcessingTimeMs:F0}";

            try
            {
                lock (_performanceLogLock)
                {
                    File.AppendAllText(_performanceLogFilePath, logEntry + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                // パフォーマンスログの書き込み失敗は通常のログに記録
                LogMessage($"パフォーマンスデータの記録に失敗: {ex.Message}");
            }
        }


        /// <summary>
        /// メッセージをログに記録し、オプションでステータスラベルに表示
        /// </summary>
        public void LogMessage(string message, bool showInStatus = false)
        {
            if (!_isLoggingEnabled) return;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            // スレッドIDを追加
            string logEntry = $"[{timestamp}][Thread:{Thread.CurrentThread.ManagedThreadId:D2}] {message}";


            try
            {
                lock (_logLock) // ログファイルへのアクセスを同期
                {
                    // ログファイルに追記
                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine, Encoding.UTF8);
                }

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
                // ログ出力自体が失敗した場合、ステータスバーだけに表示（またはデバッグ出力）
                Debug.WriteLine($"ログの記録に失敗: {ex.Message}");
                UpdateStatus($"ログの記録に失敗 (詳細はデバッグ出力参照)");
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
            LogMessage($"プロセッサ数: {Environment.ProcessorCount}");
            LogMessage($"論理プロセッサ数: {Environment.ProcessorCount}"); // より正確には `Environment.ProcessorCount`
            LogMessage($"実行ファイルパス: {Assembly.GetExecutingAssembly().Location}");


            // Officeのバージョン情報取得を試行（Late Binding方式）
            try
            {
                // Excelアプリケーションを作成
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType != null)
                {
                    object excelApp = null;
                    try
                    {
                        excelApp = Activator.CreateInstance(excelType);
                        if (excelApp != null)
                        {
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
                        }
                    }
                    finally
                    {
                        // COMの早期解放
                        if (excelApp != null)
                        {
                            try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); } catch (Exception comEx) { LogMessage($"Excel COMオブジェクト解放エラー: {comEx.Message}"); }
                        }
                    }
                }
                else
                {
                    LogMessage("Excel.ApplicationのCOMタイプが見つかりません。Excelがインストールされていないか、登録に問題がある可能性があります。");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Excelバージョン取得エラー: {ex.GetType().Name} - {ex.Message}");
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
                try
                {
                    _statusLabel.Invoke(new Action(() => _statusLabel.Text = message));
                }
                catch (ObjectDisposedException)
                {
                    // フォームが閉じられた後に呼び出された場合など
                    Debug.WriteLine($"UpdateStatus failed: Label disposed. Message: {message}");
                }
            }
            else
            {
                _statusLabel.Text = message;
            }
        }

        /// <summary>
        /// TSVフィールドのエスケープ処理 (LogPerformanceDataで使用)
        /// </summary>
        private string EscapeTsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            // タブ、改行、ダブルクォートが含まれる場合はダブルクォートで囲み、中のダブルクォートは2つにする
            if (field.Contains('\t') || field.Contains('\n') || field.Contains('\r') || field.Contains('"'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }
    }
}