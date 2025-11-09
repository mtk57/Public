using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpEmbeddingService
    {
        private const string FailureComment = "//★ メソッド特定失敗";

        public TagJumpEmbeddingResult Execute(string methodListPath, string startSourceFilePath,
            string startMethodText, string tagPrefix, TagJumpEmbeddingMode mode = TagJumpEmbeddingMode.MethodSignature,
            IProgress<TagJumpProgressInfo> progress = null, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrWhiteSpace(methodListPath))
            {
                throw new ArgumentException("メソッドリストのパスを入力してください。", nameof(methodListPath));
            }

            if (string.IsNullOrWhiteSpace(startSourceFilePath))
            {
                throw new ArgumentException("開始ソースファイルのパスを入力してください。", nameof(startSourceFilePath));
            }

            var normalizedMethodListPath = NormalizePath(methodListPath);
            if (!File.Exists(normalizedMethodListPath))
            {
                throw new FileNotFoundException("メソッドリストファイルが見つかりません。", normalizedMethodListPath);
            }

            var normalizedStartSourcePath = NormalizePath(startSourceFilePath);
            if (!File.Exists(normalizedStartSourcePath))
            {
                throw new FileNotFoundException("開始ソースファイルが見つかりません。", normalizedStartSourcePath);
            }

            if (!string.Equals(Path.GetExtension(normalizedStartSourcePath), ".java", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Javaファイル（*.java）のみが対象です。");
            }

            var entries = LoadMethodList(normalizedMethodListPath);
            if (entries.Count == 0)
            {
                throw new InvalidOperationException("メソッドリストにメソッドが登録されていません。");
            }
            ValidateRowNumberRequirement(entries, mode);

            var index = new MethodListIndex(entries);
            var startMethod = ResolveStartMethod(index, normalizedStartSourcePath, startMethodText);
            if (startMethod == null)
            {
                return new TagJumpEmbeddingResult(0, 0, Array.Empty<string>(), 0,
                    Array.Empty<TagJumpFailureDetail>());
            }
            var prefix = tagPrefix ?? string.Empty;

            var visited = new HashSet<string>(StringComparer.Ordinal);
            var analysisCache = new Dictionary<string, JavaFileStructure>(StringComparer.OrdinalIgnoreCase);
            var modifications = new Dictionary<string, List<TagInsertion>>(StringComparer.OrdinalIgnoreCase);
            var failureDetails = new List<TagJumpFailureDetail>();
            var updatedCallCount = 0;
            var processedMethodCount = 0;

            var stack = new Stack<MethodListEntry>();
            stack.Push(startMethod);

            while (stack.Count > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var current = stack.Pop();
                var methodKey = CreateMethodKey(current);
                if (!visited.Add(methodKey))
                {
                    continue;
                }

                processedMethodCount++;
                progress?.Report(new TagJumpProgressInfo("Processing", processedMethodCount, 0));

                var structure = GetOrAnalyzeStructure(current.NormalizedFilePath, analysisCache);
                var methodStructure = structure.FindMethodBySignature(current.Detail.MethodSignature);
                if (methodStructure == null)
                {
                    var failureDetail = CreateMissingMethodFailureDetail(current);
                    if (failureDetail != null)
                    {
                        failureDetails.Add(failureDetail);
                    }

                    continue;
                }

                foreach (var call in methodStructure.Calls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var calleeEntry = ProcessMethodCall(current, structure, call, index, modifications,
                        detail => failureDetails.Add(detail), prefix, normalizedMethodListPath, mode,
                        () => updatedCallCount++, cancellationToken);
                    if (calleeEntry == null)
                    {
                        continue;
                    }

                    var calleeKey = CreateMethodKey(calleeEntry);
                    if (!visited.Contains(calleeKey))
                    {
                        stack.Push(calleeEntry);
                    }
                }
            }

            return BuildResultFromModifications(modifications, analysisCache, updatedCallCount, failureDetails);
        }

        public TagJumpEmbeddingResult ExecuteAll(string methodListPath, string sourceRootDirectory,
            string tagPrefix, TagJumpEmbeddingMode mode = TagJumpEmbeddingMode.MethodSignature,
            IProgress<TagJumpProgressInfo> progress = null, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (string.IsNullOrWhiteSpace(methodListPath))
            {
                throw new ArgumentException("メソッドリストのパスを入力してください。", nameof(methodListPath));
            }

            if (string.IsNullOrWhiteSpace(sourceRootDirectory))
            {
                throw new ArgumentException("ソースルートフォルダのパスを入力してください。", nameof(sourceRootDirectory));
            }

            var normalizedMethodListPath = NormalizePath(methodListPath);
            if (!File.Exists(normalizedMethodListPath))
            {
                throw new FileNotFoundException("メソッドリストファイルが見つかりません。", normalizedMethodListPath);
            }

            var normalizedSourceRoot = NormalizePath(sourceRootDirectory);
            if (!Directory.Exists(normalizedSourceRoot))
            {
                throw new DirectoryNotFoundException($"ソースルートフォルダが見つかりません: {normalizedSourceRoot}");
            }

            var entries = LoadMethodList(normalizedMethodListPath);
            if (entries.Count == 0)
            {
                throw new InvalidOperationException("メソッドリストにメソッドが登録されていません。");
            }
            ValidateRowNumberRequirement(entries, mode);

            var index = new MethodListIndex(entries);
            var prefix = tagPrefix ?? string.Empty;

            var analysisCache = new ConcurrentDictionary<string, JavaFileStructure>(StringComparer.OrdinalIgnoreCase);
            var modifications =
                new ConcurrentDictionary<string, List<TagInsertion>>(StringComparer.OrdinalIgnoreCase);
            var failureDetails = new ConcurrentBag<TagJumpFailureDetail>();
            var processedMethods = new ConcurrentDictionary<string, byte>(StringComparer.Ordinal);
            var updatedCallCount = 0;

            var targets = new List<FileProcessingTarget>();
            foreach (var filePath in EnumerateJavaFiles(normalizedSourceRoot))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var normalizedFilePath = NormalizePath(filePath);
                var methodEntries = index.FindByFile(normalizedFilePath);
                if (methodEntries.Count == 0)
                {
                    continue;
                }

                targets.Add(new FileProcessingTarget(normalizedFilePath, methodEntries));
            }

            if (targets.Count == 0)
            {
                return new TagJumpEmbeddingResult(0, 0, Array.Empty<string>(), 0,
                    Array.Empty<TagJumpFailureDetail>());
            }

            var totalMethodTargetCount = 0;
            foreach (var target in targets)
            {
                totalMethodTargetCount += target.MethodEntries.Count;
            }

            var processedCount = 0;
            var processedFileCount = 0;
            var totalFileCount = targets.Count;
            var parallelOptions = new ParallelOptions
            {
                CancellationToken = cancellationToken
            };

            progress?.Report(new TagJumpProgressInfo("Processing", 0, totalMethodTargetCount, 0, totalFileCount));

            Parallel.ForEach(targets, parallelOptions, target =>
            {
                try
                {
                    var structure = GetOrAnalyzeStructure(target.FilePath, analysisCache);
                    foreach (var methodEntry in target.MethodEntries)
                    {
                        parallelOptions.CancellationToken.ThrowIfCancellationRequested();

                        var methodKey = CreateMethodKey(methodEntry);
                        if (!processedMethods.TryAdd(methodKey, 0))
                        {
                            continue;
                        }

                    var methodStructure = structure.FindMethodBySignature(methodEntry.Detail.MethodSignature);
                    if (methodStructure == null)
                    {
                        var missingDetail = CreateMissingMethodFailureDetail(methodEntry);
                        if (missingDetail != null)
                        {
                            failureDetails.Add(missingDetail);
                        }

                        continue;
                    }

                        foreach (var call in methodStructure.Calls)
                        {
                            parallelOptions.CancellationToken.ThrowIfCancellationRequested();
                            ProcessMethodCall(methodEntry, structure, call, index, modifications,
                                detail => failureDetails.Add(detail), prefix, normalizedMethodListPath, mode,
                                () => Interlocked.Increment(ref updatedCallCount), parallelOptions.CancellationToken);
                        }

                        var current = Interlocked.Increment(ref processedCount);
                        var currentFile = Volatile.Read(ref processedFileCount);
                        progress?.Report(new TagJumpProgressInfo("Processing", current, totalMethodTargetCount,
                            currentFile, totalFileCount));
                    }
                }
                finally
                {
                    var fileCount = Interlocked.Increment(ref processedFileCount);
                    var current = Volatile.Read(ref processedCount);
                    progress?.Report(new TagJumpProgressInfo("Processing", current, totalMethodTargetCount,
                        fileCount, totalFileCount));
                }
            });

            return BuildResultFromModifications(modifications, analysisCache, updatedCallCount, failureDetails);
        }

        private static MethodListEntry ProcessMethodCall(MethodListEntry current, JavaFileStructure structure,
            JavaMethodCallStructure call, MethodListIndex index,
            IDictionary<string, List<TagInsertion>> modifications,
            Action<TagJumpFailureDetail> failureRecorder, string prefix, string methodListPath,
            TagJumpEmbeddingMode mode, Action onInsertion, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();

            MethodListEntry calleeEntry;
            try
            {
                calleeEntry = ResolveCallee(index, current, call);
            }
            catch (MethodAmbiguityException ex)
            {
                failureRecorder?.Invoke(CreateFailureDetail(current, structure.FilePath, call,
                    "メソッドリストに同名メソッドが複数存在するため特定できませんでした。", ex.Candidates,
                    methodListPath));
                var failureInsertion = BuildFailureInsertion(structure.OriginalText, call, prefix);
                if (failureInsertion != null)
                {
                    AddInsertion(modifications, structure.FilePath, failureInsertion);
                    onInsertion?.Invoke();
                }

                return null;
            }
            catch (InvalidOperationException ex)
            {
                failureRecorder?.Invoke(CreateFailureDetail(current, structure.FilePath, call, ex.Message));
                return null;
            }

            if (calleeEntry == null)
            {
                return null;
            }

            var replacement = CreateTagReplacement(calleeEntry, prefix, methodListPath, mode);
            if (string.IsNullOrEmpty(replacement))
            {
                return null;
            }

            var insertion = BuildInsertion(structure.OriginalText, call, replacement, prefix);
            if (insertion != null)
            {
                AddInsertion(modifications, structure.FilePath, insertion);
                onInsertion?.Invoke();
            }

            return calleeEntry;
        }

        private static TagJumpEmbeddingResult BuildResultFromModifications(
            IDictionary<string, List<TagInsertion>> modifications,
            IDictionary<string, JavaFileStructure> analysisCache, int updatedCallCount,
            IEnumerable<TagJumpFailureDetail> failureDetails)
        {
            var updatedFiles = new List<string>();
            if (modifications != null)
            {
                foreach (var pair in modifications)
                {
                    var filePath = pair.Key;
                    var operations = pair.Value;
                    if (operations == null || operations.Count == 0)
                    {
                        continue;
                    }

                    if (!analysisCache.TryGetValue(filePath, out var structure))
                    {
                        structure = GetOrAnalyzeStructure(filePath, analysisCache);
                    }

                    var newText = ApplyInsertions(structure.OriginalText, operations);
                    var encoding = structure.Encoding ?? new UTF8Encoding(false);
                    File.WriteAllText(filePath, newText, encoding);
                    updatedFiles.Add(filePath);
                }
            }

            var failures = failureDetails != null
                ? new List<TagJumpFailureDetail>(failureDetails)
                : new List<TagJumpFailureDetail>();
            return new TagJumpEmbeddingResult(updatedFiles.Count, updatedCallCount, updatedFiles,
                failures.Count, failures);
        }

        private static IEnumerable<string> EnumerateJavaFiles(string rootDirectory)
        {
            if (string.IsNullOrEmpty(rootDirectory))
            {
                yield break;
            }

            var stack = new Stack<string>();
            stack.Push(rootDirectory);

            while (stack.Count > 0)
            {
                var current = stack.Pop();
                string[] files;
                try
                {
                    files = Directory.GetFiles(current, "*.java");
                }
                catch
                {
                    files = Array.Empty<string>();
                }

                foreach (var file in files)
                {
                    yield return file;
                }

                string[] subDirectories;
                try
                {
                    subDirectories = Directory.GetDirectories(current);
                }
                catch
                {
                    subDirectories = Array.Empty<string>();
                }

                foreach (var directory in subDirectories)
                {
                    stack.Push(directory);
                }
            }
        }

        private static JavaFileStructure GetOrAnalyzeStructure(string filePath,
            IDictionary<string, JavaFileStructure> cache)
        {
            if (!cache.TryGetValue(filePath, out var structure))
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("解析対象のソースファイルが見つかりません。", filePath);
                }

                structure = JavaMethodCallAnalyzer.ExtractMethodStructures(filePath);
                cache[filePath] = structure;
            }

            return structure;
        }

        private static void AddInsertion(IDictionary<string, List<TagInsertion>> map, string filePath,
            TagInsertion insertion)
        {
            if (insertion == null)
            {
                return;
            }

            if (map is ConcurrentDictionary<string, List<TagInsertion>> concurrentDictionary)
            {
                var insertionList = concurrentDictionary.GetOrAdd(filePath, _ => new List<TagInsertion>());
                lock (insertionList)
                {
                    insertionList.Add(insertion);
                }

                return;
            }

            if (!map.TryGetValue(filePath, out var list))
            {
                list = new List<TagInsertion>();
                map[filePath] = list;
            }

            list.Add(insertion);
        }

        private static string CreateTagReplacement(MethodListEntry calleeEntry, string prefix, string methodListPath, TagJumpEmbeddingMode mode)
        {
            if (calleeEntry == null)
            {
                return string.Empty;
            }

            string tagContent;
            if (mode == TagJumpEmbeddingMode.RowNumber)
            {
                var lineNumber = calleeEntry.Detail?.LineNumber ?? -1;
                if (lineNumber <= 0)
                {
                    return string.Empty;
                }

                tagContent = string.Concat(
                    calleeEntry.Detail.FilePath,
                    "\t",
                    lineNumber.ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                var methodListSegment = methodListPath ?? string.Empty;
                if (string.IsNullOrEmpty(methodListSegment))
                {
                    tagContent = string.Concat(calleeEntry.Detail.FilePath, "\t", calleeEntry.Detail.MethodSignature);
                }
                else
                {
                    tagContent = string.Concat(calleeEntry.Detail.FilePath, "\t", calleeEntry.Detail.MethodSignature, "\t", methodListSegment);
                }
            }

            return (prefix ?? string.Empty) + tagContent;
        }

        private static TagInsertion BuildInsertion(string originalText, JavaMethodCallStructure call,
            string replacement, string prefix)
        {
            if (call == null || string.IsNullOrEmpty(originalText) || string.IsNullOrEmpty(replacement))
            {
                return null;
            }

            var tailStart = call.CallEndIndex + 1;
            if (tailStart < 0)
            {
                tailStart = 0;
            }
            else if (tailStart > originalText.Length)
            {
                tailStart = originalText.Length;
            }

            var lineEnd = FindLineEndExclusive(originalText, call.CallEndIndex);
            if (lineEnd < tailStart)
            {
                lineEnd = tailStart;
            }

            var tailLength = lineEnd - tailStart;
            var tail = tailLength > 0 ? originalText.Substring(tailStart, tailLength) : string.Empty;
            var cleanedTail = RemoveExistingTagFromTail(originalText, tailStart, lineEnd, prefix);
            var newTail = AppendReplacementToTail(cleanedTail, replacement);

            if (string.Equals(tail, newTail, StringComparison.Ordinal))
            {
                return null;
            }

            return new TagInsertion(tailStart, tailLength, newTail);
        }

        private static string RemoveExistingTagFromTail(string originalText, int tailStart, int tailEnd,
            string prefix)
        {
            if (string.IsNullOrEmpty(originalText) || tailStart >= tailEnd)
            {
                return string.Empty;
            }

            var tailLength = tailEnd - tailStart;
            var tail = originalText.Substring(tailStart, tailLength);
            var segment = TagJumpSyntaxHelper.FindFirstPrefixSegmentInRange(originalText, tailStart, tailEnd, prefix);
            if (!segment.HasValue)
            {
                return tail;
            }

            var relativeStart = segment.Start - tailStart;
            if (relativeStart < 0 || relativeStart > tail.Length)
            {
                return tail;
            }

            var removalEnd = relativeStart + segment.Length;
            if (removalEnd > tail.Length)
            {
                removalEnd = tail.Length;
            }

            var before = tail.Substring(0, relativeStart);
            var after = tail.Substring(removalEnd);
            if (before.Length > 0 && after.Length > 0 &&
                char.IsWhiteSpace(before[before.Length - 1]) && char.IsWhiteSpace(after[0]))
            {
                before = before.Substring(0, before.Length - 1);
            }
            return before + after;
        }

        private static string AppendReplacementToTail(string tail, string replacement)
        {
            var baseTail = tail ?? string.Empty;
            if (string.IsNullOrEmpty(replacement))
            {
                return baseTail;
            }

            var builder = new StringBuilder(baseTail.Length + replacement.Length + 1);
            builder.Append(baseTail);
            if (builder.Length > 0)
            {
                var last = builder[builder.Length - 1];
                if (!char.IsWhiteSpace(last))
                {
                    builder.Append(' ');
                }
            }

            builder.Append(replacement);
            return builder.ToString();
        }

        private static int FindLineEndExclusive(string text, int index)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var limit = text.Length;
            var current = index;
            if (current < 0)
            {
                current = 0;
            }
            else if (current > limit)
            {
                current = limit;
            }

            while (current < limit)
            {
                var ch = text[current];
                if (ch == '\r' || ch == '\n')
                {
                    break;
                }

                current++;
            }

            return current;
        }

        private static bool StartsWith(string text, int index, string value)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value))
            {
                return false;
            }

            if (index < 0 || index + value.Length > text.Length)
            {
                return false;
            }

            return string.Compare(text, index, value, 0, value.Length, StringComparison.Ordinal) == 0;
        }

        private static string ApplyInsertions(string originalText, List<TagInsertion> insertions)
        {
            if (insertions == null || insertions.Count == 0)
            {
                return originalText;
            }

            var builder = new StringBuilder(originalText);
            insertions.Sort((x, y) => y.Index.CompareTo(x.Index));

            foreach (var insertion in insertions)
            {
                if (insertion.Index < 0 || insertion.Index > builder.Length)
                {
                    continue;
                }

                if (insertion.Length > 0)
                {
                    var removable = Math.Min(insertion.Length, builder.Length - insertion.Index);
                    if (removable > 0)
                    {
                        builder.Remove(insertion.Index, removable);
                    }
                }

                builder.Insert(insertion.Index, insertion.Text);
            }

            return builder.ToString();
        }

        private static TagInsertion BuildFailureInsertion(string originalText, JavaMethodCallStructure call, string configuredPrefix)
        {
            if (call == null || string.IsNullOrEmpty(originalText))
            {
                return null;
            }

            var insertBase = call.CallEndIndex + 1;
            if (insertBase < 0)
            {
                insertBase = 0;
            }

            if (insertBase > originalText.Length)
            {
                insertBase = originalText.Length;
            }

            var lineEnd = insertBase;
            while (lineEnd < originalText.Length)
            {
                var ch = originalText[lineEnd];
                if (ch == '\r' || ch == '\n')
                {
                    break;
                }

                lineEnd++;
            }

            var segmentStart = insertBase;
            while (segmentStart < lineEnd && (originalText[segmentStart] == ' ' || originalText[segmentStart] == '\t'))
            {
                segmentStart++;
            }

            if (segmentStart < lineEnd && originalText[segmentStart] == ';')
            {
                segmentStart++;
                while (segmentStart < lineEnd && (originalText[segmentStart] == ' ' || originalText[segmentStart] == '\t'))
                {
                    segmentStart++;
                }
            }

            var existingLength = lineEnd - segmentStart;
            if (existingLength >= 0)
            {
                var prefixSegment =
                    TagJumpSyntaxHelper.FindExistingPrefixSegment(originalText, segmentStart, configuredPrefix);
                if (prefixSegment.HasValue)
                {
                    return null;
                }
            }

            if (existingLength > 0)
            {
                var existing = originalText.Substring(segmentStart, existingLength);
                if (existing.IndexOf(FailureComment, StringComparison.Ordinal) >= 0)
                {
                    return null;
                }
            }

            var needSpace = lineEnd > 0 && originalText[lineEnd - 1] != ' ' && originalText[lineEnd - 1] != '\t';
            var text = (needSpace ? " " : string.Empty) + FailureComment;
            return new TagInsertion(lineEnd, 0, text);
        }

        private static TagJumpFailureDetail CreateFailureDetail(MethodListEntry currentMethod, string filePath,
            JavaMethodCallStructure call, string reason,
            IReadOnlyList<MethodListEntry> candidates = null, string methodListPath = null)
        {
            var callerSignature = currentMethod?.Detail?.MethodSignature ?? string.Empty;
            var callExpression = BuildCallExpression(call);
            var lineNumber = call?.LineNumber ?? 0;
            var reasonText = BuildReasonText(reason, candidates, methodListPath);
            return new TagJumpFailureDetail(filePath, lineNumber, callerSignature, callExpression, reasonText);
        }

        private static TagJumpFailureDetail CreateMissingMethodFailureDetail(MethodListEntry entry)
        {
            if (entry?.Detail == null)
            {
                return null;
            }

            var filePath = entry.Detail.FilePath ?? entry.NormalizedFilePath ?? string.Empty;
            var signature = entry.Detail.MethodSignature ?? entry.MethodName ?? string.Empty;
            var reasonBuilder = new StringBuilder("ソース内にメソッドシグネチャが見つかりませんでした。");
            if (!string.IsNullOrEmpty(signature))
            {
                reasonBuilder.Append(" 対象: ");
                reasonBuilder.Append(signature);
            }

            return new TagJumpFailureDetail(filePath, 0, signature, string.Empty, reasonBuilder.ToString());
        }

        private static string BuildCallExpression(JavaMethodCallStructure call)
        {
            if (call == null)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            if (call.HasExplicitCallee && !string.IsNullOrEmpty(call.CalleeIdentifier))
            {
                builder.Append(call.CalleeIdentifier);
                builder.Append('.');
            }

            builder.Append(call.MethodName);
            var arguments = call.ArgumentsText;
            if (!string.IsNullOrEmpty(arguments))
            {
                builder.Append(arguments);
            }
            else
            {
                builder.Append("()");
            }

            return builder.ToString();
        }

        private static string BuildReasonText(string reason,
            IReadOnlyList<MethodListEntry> candidates, string methodListPath)
        {
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(reason))
            {
                builder.Append(reason);
            }

            if (candidates != null && candidates.Count > 0)
            {
                if (builder.Length > 0)
                {
                    builder.AppendLine();
                }

                builder.AppendLine("候補一覧:");
                foreach (var candidate in candidates)
                {
                    if (candidate == null)
                    {
                        continue;
                    }

                    builder.Append(" - //@ ");
                    builder.Append(candidate.Detail?.FilePath ?? string.Empty);
                    builder.Append('\t');
                    builder.Append(candidate.Detail?.MethodSignature ?? string.Empty);
                    if (!string.IsNullOrEmpty(methodListPath))
                    {
                        builder.Append('\t');
                        builder.Append(methodListPath);
                    }

                    builder.AppendLine();
                }
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private static MethodListEntry ResolveStartMethod(MethodListIndex index, string normalizedStartSourcePath,
            string startMethodText)
        {
            var candidates = index.FindByFile(normalizedStartSourcePath);
            if (candidates.Count == 0)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(startMethodText))
            {
                return null;
            }

            var trimmed = startMethodText.Trim();
            if (trimmed.IndexOf('(') >= 0)
            {
                foreach (var entry in candidates)
                {
                    if (string.Equals(entry.Detail.MethodSignature, trimmed, StringComparison.Ordinal))
                    {
                        return entry;
                    }
                }

                return null;
            }

            var matched = new List<MethodListEntry>();
            foreach (var entry in candidates)
            {
                if (string.Equals(entry.MethodName, trimmed, StringComparison.Ordinal))
                {
                    matched.Add(entry);
                }
            }

            if (matched.Count == 0)
            {
                return null;
            }

            if (matched.Count > 1)
            {
                throw new InvalidOperationException("同名のメソッドが複数見つかりました。メソッドリストのシグネチャを指定してください。");
            }

            return matched[0];
        }

        private static MethodListEntry ResolveCallee(MethodListIndex index, MethodListEntry currentMethod,
            JavaMethodCallStructure call)
        {
            if (call == null)
            {
                return null;
            }

            var methodName = call.MethodName;
            if (string.IsNullOrEmpty(methodName))
            {
                return null;
            }

            List<MethodListEntry> candidates;

            if (!call.HasExplicitCallee)
            {
                candidates = new List<MethodListEntry>(
                    index.FindByFileAndMethodName(currentMethod.NormalizedFilePath, methodName));
            }
            else
            {
                candidates = new List<MethodListEntry>(
                    index.FindByClassAndMethodName(call.CalleeIdentifier, methodName));
            }

            candidates = FilterByArgumentCount(candidates, call.ArgumentCount);
            if (candidates.Count == 1)
            {
                return candidates[0];
            }

            if (candidates.Count > 1)
            {
                throw BuildAmbiguousCallException(currentMethod, call, candidates);
            }

            candidates = FilterByArgumentCount(
                new List<MethodListEntry>(index.FindByMethodName(methodName)), call.ArgumentCount);

            if (candidates.Count == 0)
            {
                return null;
            }

            if (candidates.Count > 1)
            {
                throw BuildAmbiguousCallException(currentMethod, call, candidates);
            }

            return candidates[0];
        }

        private static List<MethodListEntry> FilterByArgumentCount(List<MethodListEntry> source, int argumentCount)
        {
            if (source == null || source.Count == 0)
            {
                return new List<MethodListEntry>();
            }

            var filtered = new List<MethodListEntry>();
            foreach (var entry in source)
            {
                if (entry.ParameterCount == argumentCount)
                {
                    filtered.Add(entry);
                }
            }

            if (filtered.Count > 0)
            {
                return filtered;
            }

            return new List<MethodListEntry>(source);
        }

        private static MethodAmbiguityException BuildAmbiguousCallException(MethodListEntry currentMethod,
            JavaMethodCallStructure call, IReadOnlyCollection<MethodListEntry> candidates)
        {
            var builder = new StringBuilder();
            builder.AppendLine("呼び出し先メソッドを一意に特定できませんでした。");
            builder.AppendLine($"ファイル: {currentMethod.Detail.FilePath}");
            builder.AppendLine($"呼び出し元: {currentMethod.Detail.MethodSignature}");
            builder.AppendLine($"呼び出し行: {call.LineNumber}");
            builder.AppendLine("候補一覧:");
            foreach (var candidate in candidates)
            {
                builder.Append("  - ");
                builder.Append(candidate.Detail.FilePath);
                builder.Append(" : ");
                builder.Append(candidate.Detail.MethodSignature);
                builder.AppendLine();
            }

            return new MethodAmbiguityException(builder.ToString(), new List<MethodListEntry>(candidates));
        }

        private static List<MethodListEntry> LoadMethodList(string methodListPath)
        {
            var lines = File.ReadAllLines(methodListPath, Encoding.UTF8);
            var results = new List<MethodListEntry>();
            if (lines.Length <= 1)
            {
                return results;
            }

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

                var filePath = NormalizePath(parts[0]);
                var packageName = parts[2].Trim();
                var className = parts[3].Trim();
                var methodSignature = parts[4].Trim();
                if (methodSignature.Length == 0)
                {
                    continue;
                }

                var lineNumber = -1;
                if (parts.Length >= 6)
                {
                    var rowText = parts[5].Trim();
                    if (rowText.Length > 0 && int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0)
                    {
                        lineNumber = parsed;
                    }
                }

                var detail = new MethodDefinitionDetail(filePath, packageName, className, methodSignature, lineNumber);
                var methodName = GetMethodNameFromSignature(methodSignature);
                var parameterCount = CountParametersFromSignature(methodSignature);
                results.Add(new MethodListEntry(detail, filePath, methodName, parameterCount));
            }

            return results;
        }

        private static void ValidateRowNumberRequirement(List<MethodListEntry> entries, TagJumpEmbeddingMode mode)
        {
            if (mode != TagJumpEmbeddingMode.RowNumber)
            {
                return;
            }

            if (entries == null || entries.Count == 0)
            {
                throw new InvalidOperationException("メソッドリストにメソッドが登録されていません。");
            }

            var missingRow = entries.Exists(entry => entry?.Detail == null || entry.Detail.LineNumber <= 0);
            if (missingRow)
            {
                throw new InvalidOperationException("行番号版メソッドリストを指定してください。（RowNum 列が不足しています）");
            }
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

        private static string GetMethodNameFromSignature(string signature)
        {
            if (string.IsNullOrEmpty(signature))
            {
                return string.Empty;
            }

            var openParen = signature.IndexOf('(');
            if (openParen < 0)
            {
                return signature.Trim();
            }

            var beforeParen = signature.Substring(0, openParen).TrimEnd();
            var lastSpace = beforeParen.LastIndexOf(' ');
            if (lastSpace < 0)
            {
                return beforeParen.Trim();
            }

            return beforeParen.Substring(lastSpace + 1).Trim();
        }

        private static int CountParametersFromSignature(string signature)
        {
            if (string.IsNullOrEmpty(signature))
            {
                return 0;
            }

            var openParen = signature.IndexOf('(');
            var closeParen = signature.LastIndexOf(')');
            if (openParen < 0 || closeParen < 0 || closeParen <= openParen)
            {
                return 0;
            }

            var content = signature.Substring(openParen + 1, closeParen - openParen - 1).Trim();
            if (content.Length == 0)
            {
                return 0;
            }

            var count = 1;
            var angleDepth = 0;
            var bracketDepth = 0;
            var braceDepth = 0;
            var parenthesesDepth = 0;

            for (var i = 0; i < content.Length; i++)
            {
                var c = content[i];
                switch (c)
                {
                    case '<':
                        angleDepth++;
                        break;
                    case '>':
                        if (angleDepth > 0)
                        {
                            angleDepth--;
                        }
                        break;
                    case '[':
                        bracketDepth++;
                        break;
                    case ']':
                        if (bracketDepth > 0)
                        {
                            bracketDepth--;
                        }
                        break;
                    case '{':
                        braceDepth++;
                        break;
                    case '}':
                        if (braceDepth > 0)
                        {
                            braceDepth--;
                        }
                        break;
                    case '(':
                        parenthesesDepth++;
                        break;
                    case ')':
                        if (parenthesesDepth > 0)
                        {
                            parenthesesDepth--;
                        }
                        break;
                    case ',':
                        if (angleDepth == 0 && bracketDepth == 0 && braceDepth == 0 && parenthesesDepth == 0)
                        {
                            count++;
                        }

                        break;
                }
            }

            return count;
        }

        private static string CreateMethodKey(MethodListEntry entry)
        {
            return $"{entry.NormalizedFilePath}|{entry.Detail.MethodSignature}";
        }

        private sealed class FileProcessingTarget
        {
            public FileProcessingTarget(string filePath, List<MethodListEntry> methodEntries)
            {
                FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
                MethodEntries = methodEntries ?? throw new ArgumentNullException(nameof(methodEntries));
            }

            public string FilePath { get; }

            public List<MethodListEntry> MethodEntries { get; }
        }

        private sealed class MethodListEntry
        {
            public MethodListEntry(MethodDefinitionDetail detail, string normalizedFilePath,
                string methodName, int parameterCount)
            {
                Detail = detail ?? throw new ArgumentNullException(nameof(detail));
                NormalizedFilePath = normalizedFilePath ?? string.Empty;
                MethodName = methodName ?? string.Empty;
                ParameterCount = parameterCount < 0 ? 0 : parameterCount;
            }

            public MethodDefinitionDetail Detail { get; }
            public string NormalizedFilePath { get; }
            public string MethodName { get; }
            public int ParameterCount { get; }
        }

        private sealed class MethodAmbiguityException : InvalidOperationException
        {
            public MethodAmbiguityException(string message, IReadOnlyList<MethodListEntry> candidates)
                : base(message)
            {
                Candidates = candidates ?? Array.Empty<MethodListEntry>();
            }

            public IReadOnlyList<MethodListEntry> Candidates { get; }
        }

        private sealed class MethodListIndex
        {
            private readonly Dictionary<string, List<MethodListEntry>> _byFilePath;
            private readonly Dictionary<string, List<MethodListEntry>> _byClassName;
            private readonly Dictionary<string, List<MethodListEntry>> _byMethodName;

            public MethodListIndex(IEnumerable<MethodListEntry> entries)
            {
                _byFilePath = new Dictionary<string, List<MethodListEntry>>(StringComparer.OrdinalIgnoreCase);
                _byClassName = new Dictionary<string, List<MethodListEntry>>(StringComparer.Ordinal);
                _byMethodName = new Dictionary<string, List<MethodListEntry>>(StringComparer.Ordinal);

                foreach (var entry in entries)
                {
                    AddToDictionary(_byFilePath, entry.NormalizedFilePath, entry);
                    AddToDictionary(_byClassName, entry.Detail.ClassName ?? string.Empty, entry);
                    AddToDictionary(_byMethodName, entry.MethodName, entry);
                }
            }

            public List<MethodListEntry> FindByFile(string normalizedFilePath)
            {
                if (normalizedFilePath != null && _byFilePath.TryGetValue(normalizedFilePath, out var list))
                {
                    return new List<MethodListEntry>(list);
                }

                return new List<MethodListEntry>();
            }

            public List<MethodListEntry> FindByFileAndMethodName(string normalizedFilePath, string methodName)
            {
                var results = new List<MethodListEntry>();
                if (normalizedFilePath == null)
                {
                    return results;
                }

                if (_byFilePath.TryGetValue(normalizedFilePath, out var list))
                {
                    foreach (var entry in list)
                    {
                        if (string.Equals(entry.MethodName, methodName, StringComparison.Ordinal))
                        {
                            results.Add(entry);
                        }
                    }
                }

                return results;
            }

            public List<MethodListEntry> FindByClassAndMethodName(string className, string methodName)
            {
                var results = new List<MethodListEntry>();
                if (className == null)
                {
                    return results;
                }

                if (_byClassName.TryGetValue(className, out var list))
                {
                    foreach (var entry in list)
                    {
                        if (string.Equals(entry.MethodName, methodName, StringComparison.Ordinal))
                        {
                            results.Add(entry);
                        }
                    }
                }

                return results;
            }

            public List<MethodListEntry> FindByMethodName(string methodName)
            {
                if (methodName != null && _byMethodName.TryGetValue(methodName, out var list))
                {
                    return new List<MethodListEntry>(list);
                }

                return new List<MethodListEntry>();
            }

            private static void AddToDictionary(IDictionary<string, List<MethodListEntry>> dictionary,
                string key, MethodListEntry entry)
            {
                if (!dictionary.TryGetValue(key, out var list))
                {
                    list = new List<MethodListEntry>();
                    dictionary[key] = list;
                }

                list.Add(entry);
            }
        }

        private sealed class TagInsertion
        {
            public TagInsertion(int index, int length, string text)
            {
                Index = index;
                Length = length < 0 ? 0 : length;
                Text = text ?? string.Empty;
            }

            public int Index { get; }
            public int Length { get; }
            public string Text { get; }
        }
    }
}
