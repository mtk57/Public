using System;
using System.IO;
using System.Text;
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

            var lineText = args[0] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(lineText))
            {
                return false;
            }

            var methodListPath = string.Empty;
            if (args.Length > 1)
            {
                methodListPath = args[1] ?? string.Empty;
            }

            RunCommandLineMode(lineText, methodListPath);
            return true;
        }

        private static void RunCommandLineMode(string lineText, string methodListPath)
        {
            string filePath = string.Empty;
            string methodSignature = string.Empty;
            string resolvedMethodListPath = string.Empty;

            try
            {
                var settings = SettingsManager.Load();
                var prefix = settings?.LastTagJumpPrefix;
                if (string.IsNullOrEmpty(prefix))
                {
                    prefix = "//@ ";
                }

                if (string.IsNullOrWhiteSpace(methodListPath))
                {
                    methodListPath = settings?.LastTagJumpMethodListPath ?? string.Empty;
                }

                if (string.IsNullOrWhiteSpace(methodListPath))
                {
                    var message = "メソッドリストのパスが指定されていません。アプリ側で設定してください。";
                    Console.Error.WriteLine(message);
                    ErrorLogger.LogError(message);
                    MessageBox.Show(message, "タグジャンプエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.ExitCode = 1;
                    return;
                }

                resolvedMethodListPath = methodListPath;
                if (!TagJumpSyntaxHelper.TryParseTagLine(lineText, prefix, out filePath, out methodSignature))
                {
                    var builder = new StringBuilder();
                    builder.AppendLine("タグジャンプ情報を解析できませんでした。");
                    builder.AppendLine($"行内容: {lineText}");
                    var message = builder.ToString().TrimEnd();
                    Console.Error.WriteLine(message);
                    ErrorLogger.LogError(message);
                    MessageBox.Show(message, "タグジャンプエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.ExitCode = 1;
                    return;
                }

                if (!File.Exists(filePath))
                {
                    var message = $"タグジャンプ対象のファイルが見つかりません。ファイル: {filePath}";
                    Console.Error.WriteLine(message);
                    ErrorLogger.LogError(message);
                    MessageBox.Show(message, "タグジャンプエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.ExitCode = 1;
                    return;
                }

                var methodDetail = TagJumpMethodLocator.FindMethod(resolvedMethodListPath, filePath, methodSignature);
                if (methodDetail == null)
                {
                    var message = $"指定されたメソッドが見つかりませんでした。シグネチャ: {methodSignature}";
                    Console.Error.WriteLine(message);
                    ErrorLogger.LogError(message);
                    MessageBox.Show(message, "タグジャンプエラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.ExitCode = 1;
                    return;
                }

                filePath = methodDetail.FilePath;
                var lineNumber = methodDetail.LineNumber > 0 ? methodDetail.LineNumber : 1;
                SakuraTagJumpLauncher.Launch(filePath, lineNumber);
                Environment.ExitCode = 0;
            }
            catch (JavaParseException ex)
            {
                var builder = new StringBuilder();
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
                MessageBox.Show(message, "解析エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.ExitCode = 1;
            }
            catch (Exception ex)
            {
                var builder = new StringBuilder();
                builder.AppendLine("コマンドライン処理中にエラーが発生しました。");
                if (!string.IsNullOrEmpty(resolvedMethodListPath))
                {
                    builder.AppendLine($"メソッドリスト: {resolvedMethodListPath}");
                }
                if (!string.IsNullOrEmpty(filePath))
                {
                    builder.AppendLine($"ファイル: {filePath}");
                }
                if (!string.IsNullOrEmpty(methodSignature))
                {
                    builder.AppendLine($"メソッド: {methodSignature}");
                }

                builder.AppendLine(ex.Message);
                var message = builder.ToString().TrimEnd();
                Console.Error.WriteLine(message);
                ErrorLogger.LogException(ex);
                MessageBox.Show(message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Environment.ExitCode = 1;
            }
        }

    }
}
