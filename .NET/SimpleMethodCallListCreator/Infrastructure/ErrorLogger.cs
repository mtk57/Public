using System;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class ErrorLogger
    {
        private const string LogFileName = "error.log";

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

                File.AppendAllText(path, content);
            }
            catch
            {
                // ログ出力に失敗してもアプリケーションの挙動には影響させない
            }
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
