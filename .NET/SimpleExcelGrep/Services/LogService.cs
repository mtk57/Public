﻿using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// アプリケーションのログ記録を担当するサービス
    /// </summary>
    internal class LogService
    {
        private readonly string _logFilePath;
        private readonly bool _isLoggingEnabled;
        private readonly Label _statusLabel;

        /// <summary>
        /// LogServiceのコンストラクタ
        /// </summary>
        /// <param name="statusLabel">ステータス表示に使用するUIラベル</param>
        /// <param name="isLoggingEnabled">ログ記録を有効にするかどうか</param>
        public LogService(Label statusLabel, bool isLoggingEnabled = true)
        {
            _statusLabel = statusLabel;
            _isLoggingEnabled = isLoggingEnabled;
            _logFilePath = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "SimpleExcelGrep_log.txt");
        }

        /// <summary>
        /// メッセージをログに記録し、オプションでステータスラベルに表示
        /// </summary>
        public void LogMessage(string message, bool showInStatus = false)
        {
            if (!_isLoggingEnabled) return;

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
                    object excelApp = null;
                    object version = null;
                    object build = null;
                    
                    try
                    {
                        excelApp = Activator.CreateInstance(excelType);

                        // バージョン情報を取得
                        version = excelType.InvokeMember("Version",
                            System.Reflection.BindingFlags.GetProperty,
                            null, excelApp, null);
                        LogMessage($"Excel バージョン: {version}");

                        // ビルド情報を取得
                        build = excelType.InvokeMember("Build",
                            System.Reflection.BindingFlags.GetProperty,
                            null, excelApp, null);
                        LogMessage($"Excel ビルド: {build}");
                    }
                    finally
                    {
                        // 修正: COMオブジェクトの解放
                        if (build != null && System.Runtime.InteropServices.Marshal.IsComObject(build))
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(build);
                            build = null;
                        }
                        
                        if (version != null && System.Runtime.InteropServices.Marshal.IsComObject(version))
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(version);
                            version = null;
                        }
                        
                        if (excelApp != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                            excelApp = null;
                        }
                    }
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

        /// <summary>
        /// 現在のメモリ使用量をログに記録
        /// </summary>
        public void LogMemoryUsage(string contextInfo)
        {
            // GC前のメモリ使用量を取得
            long memoryBeforeCollect = GC.GetTotalMemory(false);
            double memoryBeforeMB = memoryBeforeCollect / (1024.0 * 1024.0);
    
            // GCを実行
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
    
            // GC後のメモリ使用量を取得
            long memoryAfterCollect = GC.GetTotalMemory(true);
            double memoryAfterMB = memoryAfterCollect / (1024.0 * 1024.0);
    
            // プロセスの総メモリ使用量を取得 (より正確)
            long processMemory = System.Diagnostics.Process.GetCurrentProcess().WorkingSet64;
            double processMemoryMB = processMemory / (1024.0 * 1024.0);
    
            // ログに記録
            LogMessage($"メモリ使用量 ({contextInfo}): GC前={memoryBeforeMB:F2}MB, " +
                       $"GC後={memoryAfterMB:F2}MB, プロセス合計={processMemoryMB:F2}MB");
    
            // メモリリークの可能性がある場合に警告
            if (memoryAfterMB > 200 || (memoryBeforeMB - memoryAfterMB) > 50)
            {
                LogMessage($"警告: メモリ使用量が高い、または解放されていないメモリがある可能性があります。GC差分={(memoryBeforeMB - memoryAfterMB):F2}MB");
            }
        }

        /// <summary>
        /// メモリリークの可能性がある場所を診断
        /// </summary>
        public void DiagnoseMemoryLeak(string methodName, Action action)
        {
            LogMessage($"メモリリーク診断開始: {methodName}");
    
            // 実行前のメモリ使用量を取得
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long memoryBefore = GC.GetTotalMemory(true);
    
            try
            {
                // テスト対象のコードを実行
                action();
            }
            finally
            {
                // 実行後のメモリ使用量を取得
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                long memoryAfter = GC.GetTotalMemory(true);
        
                // 差分を計算
                double diffMB = (memoryAfter - memoryBefore) / (1024.0 * 1024.0);
        
                if (diffMB > 0.5) // 0.5MB以上の差がある場合は警告
                {
                    LogMessage($"警告: {methodName} で {diffMB:F2}MB のメモリが解放されていない可能性があります。");
                }
                else
                {
                    LogMessage($"メモリリーク診断完了: {methodName}, 差分={diffMB:F2}MB");
                }
            }
        }
    }
}