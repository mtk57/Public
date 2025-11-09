using System;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class ErrorLogger
    {
        private static readonly Encoding LogEncoding = Encoding.GetEncoding("shift_jis");
        private const string WindowsNewLine = "\r\n";
        private static readonly object SyncRoot = new object();
        private static string _logFilePath;

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
                EnsureShiftJisEncoding(path);
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

        private static void EnsureShiftJisEncoding(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return;
            }

            try
            {
                if (!HasUtf8Bom(path))
                {
                    return;
                }

                var content = File.ReadAllText(path, Encoding.UTF8);
                var normalized = NormalizeNewLines(content);
                File.WriteAllText(path, normalized, LogEncoding);
            }
            catch
            {
                // 変換に失敗しても既存の内容を優先する
            }
        }

        private static bool HasUtf8Bom(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                if (!stream.CanRead || stream.Length < 3)
                {
                    return false;
                }

                var buffer = new byte[3];
                var read = stream.Read(buffer, 0, buffer.Length);
                if (read < 3)
                {
                    return false;
                }

                return buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF;
            }
        }

        public static string GetLogFilePath()
        {
            if (!string.IsNullOrEmpty(_logFilePath))
            {
                return _logFilePath;
            }

            lock (SyncRoot)
            {
                if (!string.IsNullOrEmpty(_logFilePath))
                {
                    return _logFilePath;
                }

                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                if (string.IsNullOrEmpty(baseDirectory))
                {
                    baseDirectory = Environment.CurrentDirectory;
                }

                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var fileName = $"Error_{timestamp}.log";
                _logFilePath = Path.Combine(baseDirectory, fileName);
                return _logFilePath;
            }
        }
    }
}
