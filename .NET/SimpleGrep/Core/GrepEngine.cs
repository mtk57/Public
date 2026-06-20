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
            IReadOnlyList<string> grepPatterns,
            bool caseSensitive,
            bool useRegex,
            bool deriveMethod,
            bool ignoreComment,
            IProgress<int> progress,
            CancellationToken cancellationToken,
            Action<Exception> errorHandler)
        {
            var results = new ConcurrentBag<SearchResult>();
            var patterns = (grepPatterns ?? Array.Empty<string>())
                .Where(pattern => !string.IsNullOrWhiteSpace(pattern))
                .ToArray();
            if (patterns.Length == 0)
            {
                return results;
            }

            var regexOptions = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            var comparisonType = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int processedFileCount = 0;

            try
            {
                List<Regex> compiledRegexes = null;
                if (useRegex)
                {
                    compiledRegexes = patterns
                        .Select(pattern => new Regex(pattern, regexOptions | RegexOptions.Compiled))
                        .ToList();
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
                            string fileExtension = Path.GetExtension(filePath);
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
                                    if (useRegex && compiledRegexes != null && compiledRegexes.Count > 0)
                                    {
                                        foreach (var regex in compiledRegexes)
                                        {
                                            if (regex.IsMatch(evaluatedLine))
                                            {
                                                isMatch = true;
                                                break;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        foreach (var pattern in patterns)
                                        {
                                            if (evaluatedLine.IndexOf(pattern, comparisonType) >= 0)
                                            {
                                                isMatch = true;
                                                break;
                                            }
                                        }
                                    }

                                    if (isMatch)
                                    {
                                        matches.Add(new SearchResult
                                        {
                                            FilePath = filePath,
                                            FileName = fileName,
                                            FileExtension = fileExtension,
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
                    if (IsUtf8(fs))
                    {
                        return new UTF8Encoding(false);
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

        private static bool IsUtf8(Stream stream)
        {
            const int BufferSize = 4096;
            var buffer = new byte[BufferSize];
            int expectedContinuationBytes = 0;
            int bytesRead;

            while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                for (int i = 0; i < bytesRead; i++)
                {
                    byte currentByte = buffer[i];
                    if (expectedContinuationBytes > 0)
                    {
                        if (currentByte < 0x80 || currentByte > 0xBF)
                        {
                            return false;
                        }

                        expectedContinuationBytes--;
                        continue;
                    }

                    if (currentByte < 0x80)
                    {
                        continue;
                    }

                    if (currentByte < 0xC2)
                    {
                        return false;
                    }

                    if (currentByte < 0xE0)
                    {
                        expectedContinuationBytes = 1;
                    }
                    else if (currentByte < 0xF0)
                    {
                        expectedContinuationBytes = 2;
                    }
                    else if (currentByte < 0xF5)
                    {
                        expectedContinuationBytes = 3;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            return expectedContinuationBytes == 0;
        }

    }
}
