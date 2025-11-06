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
        private const int Utf8ProbeLength = 64 * 1024;

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
            var comparisonType = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int processedFileCount = 0;

            try
            {
                Regex compiledRegex = null;
                if (useRegex)
                {
                    compiledRegex = new Regex(grepPattern, regexOptions | RegexOptions.Compiled);
                }

                Parallel.ForEach(
                    filePaths,
                    new ParallelOptions { CancellationToken = cancellationToken },
                    () => new List<SearchResult>(),
                    (filePath, state, localResults) =>
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        try
                        {
                            string fileName = Path.GetFileName(filePath);
                            Encoding encoding = DetectEncoding(filePath);
                            string encodingName = GetEncodingName(encoding);
                            ICommentFilter commentFilter = ignoreComment ? CommentFilterFactory.Create(filePath) : null;
                            var matches = new List<SearchResult>();
                            List<string> linesForResolver = deriveMethod ? new List<string>() : null;

                            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
                            using (var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: false))
                            {
                                string line;
                                int lineNumber = 0;
                                while ((line = reader.ReadLine()) != null)
                                {
                                    cancellationToken.ThrowIfCancellationRequested();
                                    lineNumber++;

                                    string originalLine = line ?? string.Empty;
                                    linesForResolver?.Add(originalLine);

                                    string evaluatedLine = originalLine;
                                    if (commentFilter != null)
                                    {
                                        evaluatedLine = commentFilter.RemoveComments(originalLine) ?? string.Empty;
                                        bool commentOnlyLine = !string.IsNullOrWhiteSpace(originalLine) && string.IsNullOrWhiteSpace(evaluatedLine);
                                        if (commentOnlyLine)
                                        {
                                            continue;
                                        }
                                    }

                                    bool isMatch = false;
                                    if (useRegex && compiledRegex != null)
                                    {
                                        isMatch = compiledRegex.IsMatch(evaluatedLine);
                                    }
                                    else if (!useRegex)
                                    {
                                        isMatch = evaluatedLine.IndexOf(grepPattern, comparisonType) >= 0;
                                    }

                                    if (isMatch)
                                    {
                                        matches.Add(new SearchResult
                                        {
                                            FilePath = filePath,
                                            FileName = fileName,
                                            LineNumber = lineNumber,
                                            LineText = originalLine,
                                            EncodingName = encodingName
                                        });
                                    }
                                }
                            }

                            if (deriveMethod && matches.Count > 0 && linesForResolver != null)
                            {
                                var methodResolver = CreateMethodResolver(filePath, linesForResolver);
                                if (methodResolver != null)
                                {
                                    foreach (var match in matches)
                                    {
                                        match.MethodSignature = methodResolver.GetMethodSignature(match.LineNumber) ?? string.Empty;
                                    }
                                }
                            }

                            if (matches.Count > 0)
                            {
                                localResults.AddRange(matches);
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

                        return localResults;
                    },
                    localResults =>
                    {
                        if (localResults == null)
                        {
                            return;
                        }

                        foreach (var result in localResults)
                        {
                            results.Add(result);
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
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan))
                {
                    var bom = new byte[4];
                    int read = fs.Read(bom, 0, bom.Length);

                    if (read >= 3 && bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
                    {
                        return new UTF8Encoding(true);
                    }
                    if (read >= 2)
                    {
                        if (bom[0] == 0xff && bom[1] == 0xfe)
                        {
                            return Encoding.Unicode;
                        }
                        if (bom[0] == 0xfe && bom[1] == 0xff)
                        {
                            return Encoding.BigEndianUnicode;
                        }
                    }

                    fs.Position = 0;
                    int probeLength = (int)Math.Min(fs.Length, Utf8ProbeLength);
                    if (probeLength > 0)
                    {
                        var buffer = new byte[probeLength];
                        int totalRead = 0;
                        while (totalRead < probeLength)
                        {
                            int chunk = fs.Read(buffer, totalRead, probeLength - totalRead);
                            if (chunk == 0)
                            {
                                break;
                            }
                            totalRead += chunk;
                        }

                        if (totalRead > 0 && IsUtf8(buffer, totalRead))
                        {
                            return new UTF8Encoding(false);
                        }
                    }
                }
            }
            catch
            {
                return Encoding.Default;
            }

            return Encoding.GetEncoding("Shift_JIS");
        }

        private static IMethodSignatureResolver CreateMethodResolver(string filePath, List<string> lines)
        {
            if (lines == null || lines.Count == 0)
            {
                return null;
            }

            string extension = Path.GetExtension(filePath);
            if (string.Equals(extension, ".java", StringComparison.OrdinalIgnoreCase))
            {
                return new JavaMethodSignatureResolver(lines.ToArray());
            }
            if (string.Equals(extension, ".cs", StringComparison.OrdinalIgnoreCase))
            {
                return new CSharpMethodSignatureResolver(lines.ToArray());
            }

            return null;
        }

        private static bool IsUtf8(byte[] bytes, int length)
        {
            int i = 0;
            while (i < length)
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

                if (i + extraBytes >= length)
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
