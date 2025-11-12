using System;
using System.IO;
using System.Text;

namespace SimpleSqlAdjuster
{
    internal static class LogService
    {
        private const string LogFileName = "SimpleSqlAdjuster.log";

        public static void Log(string message)
        {
            try
            {
                var basePath = AppDomain.CurrentDomain.BaseDirectory;
                var path = Path.Combine(basePath, LogFileName);
                var line = string.Format(
                    "{0:yyyy-MM-dd HH:mm:ss.fff} {1}{2}",
                    DateTime.Now,
                    message,
                    Environment.NewLine);
                File.AppendAllText(path, line, Encoding.UTF8);
            }
            catch
            {
                // ログの書き込み失敗はアプリの動作を妨げない
            }
        }

        public static void Log(Exception exception, string message = null)
        {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(message))
            {
                builder.AppendLine(message);
            }

            builder.AppendLine(exception.ToString());
            Log(builder.ToString().TrimEnd());
        }
    }
}
