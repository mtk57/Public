using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace SimpleMethodCallListCreator
{
    public static class SakuraTagJumpLauncher
    {
        public static void Launch(string filePath, int lineNumber)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("ファイルパスが指定されていません。", nameof(filePath));
            }

            if (lineNumber <= 0)
            {
                lineNumber = 1;
            }

            var sakuraPath = FindSakuraPath();
            if (!string.IsNullOrEmpty(sakuraPath))
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = sakuraPath,
                    Arguments = $"-Y={lineNumber} \"{filePath}\"",
                    UseShellExecute = false
                });
            }
            else
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
        }

        private static string FindSakuraPath()
        {
            try
            {
                var configPath = ConfigurationManager.AppSettings["SakuraEditorPath"];
                if (!string.IsNullOrEmpty(configPath) && File.Exists(configPath))
                {
                    return configPath;
                }
            }
            catch (ConfigurationErrorsException ex)
            {
                ErrorLogger.LogException(ex);
            }

            var programFilesPaths = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "sakura", "sakura.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "sakura", "sakura.exe")
            };

            foreach (var candidate in programFilesPaths)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            var pathEnv = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(pathEnv))
            {
                foreach (var part in pathEnv.Split(Path.PathSeparator))
                {
                    if (string.IsNullOrEmpty(part))
                    {
                        continue;
                    }

                    var candidate = Path.Combine(part, "sakura.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }
    }
}
