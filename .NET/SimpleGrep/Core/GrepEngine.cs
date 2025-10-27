using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SimpleGrep.Core
{
    internal static class GrepEngine
    {
        public static IEnumerable<SearchResult> SearchFiles(
            string[] filePaths,
            string grepPattern,
            bool caseSensitive,
            bool useRegex,
            bool deriveMethod,
            bool ignoreComment,
            IProgress<int> progress,
            CancellationToken cancellationToken,
            Action<Exception> errorHandler)
        {
            var results = new ConcurrentBag<SearchResult>();
            var regexOptions = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            int processedFileCount = 0;

            try
            {
                Parallel.ForEach(filePaths, new ParallelOptions { CancellationToken = cancellationToken }, filePath =>
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        string fileName = Path.GetFileName(filePath);
                        Encoding encoding = DetectEncoding(filePath);
                        string encodingName = GetEncodingName(encoding);

                        string[] lines;
                        using (var reader = new StreamReader(filePath, encoding))
                        {
                            var content = reader.ReadToEnd();
                            cancellationToken.ThrowIfCancellationRequested();
                            lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                        }

                        IMethodSignatureResolver methodResolver = null;
                        if (deriveMethod)
                        {
                            string extension = Path.GetExtension(filePath);
                            if (string.Equals(extension, ".java", StringComparison.OrdinalIgnoreCase))
                            {
                                methodResolver = new JavaMethodSignatureResolver(lines);
                            }
                            else if (string.Equals(extension, ".cs", StringComparison.OrdinalIgnoreCase))
                            {
                                methodResolver = new CSharpMethodSignatureResolver(lines);
                            }
                        }

                        ICommentFilter commentFilter = null;
                        if (ignoreComment)
                        {
                            commentFilter = CommentFilterFactory.Create(filePath);
                        }

                        for (int index = 0; index < lines.Length; index++)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            string line = lines[index] ?? string.Empty;
                            int lineNumber = index + 1;
                            bool isMatch = false;
                            string lineToEvaluate = line;

                            if (commentFilter != null)
                            {
                                lineToEvaluate = commentFilter.RemoveComments(lineToEvaluate) ?? string.Empty;
                                bool commentOnlyLine = !string.IsNullOrWhiteSpace(line) && string.IsNullOrWhiteSpace(lineToEvaluate);
                                if (commentOnlyLine)
                                {
                                    continue;
                                }
                            }

                            if (useRegex)
                            {
                                try
                                {
                                    isMatch = Regex.IsMatch(lineToEvaluate, grepPattern, regexOptions);
                                }
                                catch (ArgumentException)
                                {
                                    // 無効な正規表現パターンはスキップ
                                }
                            }
                            else
                            {
                                var comparisonType = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                                isMatch = lineToEvaluate.IndexOf(grepPattern, comparisonType) >= 0;
                            }

                            if (isMatch)
                            {
                                string methodSignature = methodResolver?.GetMethodSignature(lineNumber) ?? string.Empty;
                                results.Add(new SearchResult
                                {
                                    FilePath = filePath,
                                    FileName = fileName,
                                    LineNumber = lineNumber,
                                    LineText = line,
                                    MethodSignature = methodSignature,
                                    EncodingName = encodingName
                                });
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // Skip file read errors
                    }
                    finally
                    {
                        int currentCount = Interlocked.Increment(ref processedFileCount);
                        if (currentCount % 100 == 0 || currentCount == filePaths.Length)
                        {
                            progress?.Report(Math.Min(currentCount, filePaths.Length));
                        }
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // キャンセル時は例外を表面化させずに処理終了
            }
            catch (AggregateException ex) when (ex.InnerExceptions.All(e => e is OperationCanceledException))
            {
                // Parallel.ForEach が AggregateException でキャンセルを通知するケースに備える
            }
            catch (Exception ex)
            {
                errorHandler?.Invoke(ex);
            }

            return results;
        }

        private static string GetEncodingName(Encoding encoding)
        {
            if (encoding.CodePage == 932)
            {
                return "Shift_JIS";
            }

            if (encoding is UTF8Encoding)
            {
                return "UTF-8";
            }

            return encoding.WebName.ToUpper();
        }

        private static Encoding DetectEncoding(string filePath)
        {
            byte[] bom = new byte[4];
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    fs.Read(bom, 0, 4);
                }
            }
            catch
            {
                return Encoding.Default;
            }

            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
            {
                return new UTF8Encoding(true);
            }
            if (bom[0] == 0xff && bom[1] == 0xfe)
            {
                return Encoding.Unicode;
            }
            if (bom[0] == 0xfe && bom[1] == 0xff)
            {
                return Encoding.BigEndianUnicode;
            }

            byte[] fileBytes;
            try
            {
                fileBytes = File.ReadAllBytes(filePath);
            }
            catch
            {
                return Encoding.Default;
            }

            if (IsUtf8(fileBytes))
            {
                return new UTF8Encoding(false);
            }

            return Encoding.GetEncoding("Shift_JIS");
        }

        private static bool IsUtf8(byte[] bytes)
        {
            int i = 0;
            while (i < bytes.Length)
            {
                if (bytes[i] < 0x80)
                {
                    i++;
                    continue;
                }

                if (bytes[i] < 0xC2)
                {
                    return false;
                }

                int extraBytes = 0;
                if (bytes[i] < 0xE0) extraBytes = 1;
                else if (bytes[i] < 0xF0) extraBytes = 2;
                else if (bytes[i] < 0xF5) extraBytes = 3;
                else return false;

                if (i + extraBytes >= bytes.Length)
                {
                    return false;
                }

                for (int j = 1; j <= extraBytes; j++)
                {
                    if (bytes[i + j] < 0x80 || bytes[i + j] > 0xBF)
                    {
                        return false;
                    }
                }

                i += (extraBytes + 1);
            }

            return true;
        }
    }
}
