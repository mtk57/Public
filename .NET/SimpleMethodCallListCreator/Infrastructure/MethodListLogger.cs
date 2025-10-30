using System;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class MethodListLogger
    {
        private const string LogFileName = "methodlist.log";

        public static void LogInfo(string message)
        {
            WriteEntry("INFO", message);
        }

        public static void LogError(string message)
        {
            WriteEntry("ERROR", message);
        }

        public static void LogException(Exception exception)
        {
            if (exception == null)
            {
                return;
            }

            WriteRaw($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] EXCEPTION{Environment.NewLine}{exception}{Environment.NewLine}");
        }

        private static void WriteEntry(string level, string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            var builder = new StringBuilder();
            builder.Append('[');
            builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.Append("] ");
            builder.Append(level);
            builder.Append(' ');
            builder.AppendLine(message.Trim());
            WriteRaw(builder.ToString());
        }

        private static void WriteRaw(string content)
        {
            if (string.IsNullOrEmpty(content))
            {
                return;
            }

            try
            {
                var path = GetLogFilePath();
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.AppendAllText(path, content, Encoding.UTF8);
            }
            catch
            {
                // ログ出力で例外が発生してもアプリケーションの挙動に影響させない
            }
        }

        private static string GetLogFilePath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var directory = Path.Combine(appData, "SimpleMethodCallListCreator");
            return Path.Combine(directory, LogFileName);
        }
    }
}
