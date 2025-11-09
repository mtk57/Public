using System;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class ErrorLogger
    {
        private const string LogFileName = "error.log";
        private static readonly Encoding LogEncoding = Encoding.GetEncoding("shift_jis");
        private const string WindowsNewLine = "\r\n";

        public static void LogError(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            WriteEntry(message);
        }

        public static void LogException(Exception exception)
        {
            if (exception == null)
            {
                return;
            }

            var builder = new StringBuilder();
            builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Exception");
            builder.AppendLine(exception.ToString());
            WriteRaw(builder.ToString());
        }

        private static void WriteEntry(string message)
        {
            var builder = new StringBuilder();
            builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
            WriteRaw(builder.ToString());
        }

        private static void WriteRaw(string content)
        {
            try
            {
                var path = GetLogFilePath();
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var normalized = NormalizeNewLines(content);
                File.AppendAllText(path, normalized, LogEncoding);
            }
            catch
            {
                // ログ出力に失敗してもアプリケーションの挙動には影響させない
            }
        }

        private static string NormalizeNewLines(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return content;
            }

            var unified = content
                .Replace("\r\n", "\n")
                .Replace("\r", "\n");
            return unified.Replace("\n", WindowsNewLine);
        }

        private static string GetLogFilePath()
        {
            var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (string.IsNullOrEmpty(baseDirectory))
            {
                baseDirectory = Environment.CurrentDirectory;
            }

            return Path.Combine(baseDirectory, LogFileName);
        }
    }
}
