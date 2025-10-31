using System;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    internal static class TagJumpMethodLocator
    {
        public static MethodDefinitionDetail FindMethod(string methodListPath, string targetFilePath, string methodSignature)
        {
            if (string.IsNullOrWhiteSpace(methodListPath))
            {
                throw new ArgumentException("メソッドリストのパスが指定されていません。", nameof(methodListPath));
            }

            if (string.IsNullOrWhiteSpace(targetFilePath))
            {
                throw new ArgumentException("対象ファイルのパスが指定されていません。", nameof(targetFilePath));
            }

            if (string.IsNullOrWhiteSpace(methodSignature))
            {
                throw new ArgumentException("メソッドシグネチャが指定されていません。", nameof(methodSignature));
            }

            var normalizedMethodListPath = NormalizePath(methodListPath);
            if (!File.Exists(normalizedMethodListPath))
            {
                throw new FileNotFoundException("メソッドリストファイルが見つかりません。", normalizedMethodListPath);
            }

            var normalizedTargetFilePath = NormalizePath(targetFilePath);

            MethodDefinitionDetail listEntry = null;

            var lines = File.ReadAllLines(normalizedMethodListPath, Encoding.UTF8);
            if (lines.Length > 1)
            {
                for (var i = 1; i < lines.Length; i++)
                {
                    var line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    var parts = line.Split('\t');
                    if (parts.Length < 5)
                    {
                        throw new InvalidOperationException($"メソッドリストの形式が不正です。（{i + 1}行目）");
                    }

                    var listFilePath = NormalizePath(parts[0]);
                    if (!string.Equals(listFilePath, normalizedTargetFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var packageName = parts[2].Trim();
                    var className = parts[3].Trim();
                    var signaturePart = parts[4].Trim();
                    if (!string.Equals(signaturePart, methodSignature, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    listEntry = new MethodDefinitionDetail(listFilePath, packageName, className, signaturePart);
                    break;
                }
            }

            if (listEntry == null)
            {
                return null;
            }

            var definitions = JavaMethodCallAnalyzer.ExtractMethodDefinitions(listEntry.FilePath);
            foreach (var definition in definitions)
            {
                if (string.Equals(definition.MethodSignature, methodSignature, StringComparison.Ordinal))
                {
                    return definition;
                }
            }

            return new MethodDefinitionDetail(
                listEntry.FilePath,
                listEntry.PackageName,
                listEntry.ClassName,
                methodSignature);
        }

        private static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(path.Trim());
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"パスを正規化できませんでした。対象: {path}", ex);
            }
        }
    }
}
