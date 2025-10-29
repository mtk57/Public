using System;
using System.IO;
using System.Windows.Forms;

namespace SimpleMethodCallListCreator
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (TryRunCommandLine(args))
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }

        private static bool TryRunCommandLine(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                return false;
            }

            var filePath = (args.Length > 0 ? args[0] : string.Empty) ?? string.Empty;
            var methodSignature = (args.Length > 1 ? args[1] : string.Empty) ?? string.Empty;

            filePath = filePath.Trim();
            methodSignature = methodSignature.Trim();

            if (filePath.Length == 0 || methodSignature.Length == 0)
            {
                return false;
            }

            if (!File.Exists(filePath))
            {
                return false;
            }

            RunCommandLineMode(filePath, methodSignature);
            return true;
        }

        private static void RunCommandLineMode(string filePath, string methodSignature)
        {
            try
            {
                var definitions = JavaMethodCallAnalyzer.ExtractMethodDefinitions(filePath);
                var match = definitions.Find(detail =>
                    string.Equals(detail.MethodSignature, methodSignature, StringComparison.Ordinal));

                if (match == null)
                {
                    var message = $"指定されたメソッドが見つかりませんでした。シグネチャ: {methodSignature}";
                    Console.Error.WriteLine(message);
                    ErrorLogger.LogError(message);
                    Environment.ExitCode = 1;
                    return;
                }

                var lineNumber = match.LineNumber > 0 ? match.LineNumber : 1;
                SakuraTagJumpLauncher.Launch(filePath, lineNumber);
            }
            catch (JavaParseException ex)
            {
                var builder = new System.Text.StringBuilder();
                builder.AppendLine("Javaファイルの解析に失敗しました。");
                builder.AppendLine($"ファイル: {filePath}");
                builder.AppendLine($"行番号: {ex.LineNumber}");
                if (!string.IsNullOrEmpty(ex.InvalidContent))
                {
                    builder.AppendLine($"内容: {ex.InvalidContent}");
                }

                var message = builder.ToString().TrimEnd();
                Console.Error.WriteLine(message);
                ErrorLogger.LogError(message);
                Environment.ExitCode = 1;
            }
            catch (Exception ex)
            {
                var message = $"コマンドライン処理中にエラーが発生しました。{ex.Message}";
                Console.Error.WriteLine(message);
                ErrorLogger.LogException(ex);
                Environment.ExitCode = 1;
            }
        }
    }
}
