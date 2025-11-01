using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpEmbeddingService
    {
        private const string FailureComment = "//★ メソッド特定失敗";

        public TagJumpEmbeddingResult Execute(string methodListPath, string startSourceFilePath,
            string startMethodText, string tagPrefix)
        {
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

            var stack = new Stack<MethodListEntry>();
            stack.Push(startMethod);

            while (stack.Count > 0)
            {
                var current = stack.Pop();
                var methodKey = CreateMethodKey(current);
                if (!visited.Add(methodKey))
                {
                    continue;
                }

                var structure = GetOrAnalyzeStructure(current.NormalizedFilePath, analysisCache);
                var methodStructure = structure.FindMethodBySignature(current.Detail.MethodSignature);
                if (methodStructure == null)
                {
                    throw new InvalidOperationException(
                        $"ソース内にメソッドシグネチャが見つかりませんでした。ファイル: {current.Detail.FilePath}, シグネチャ: {current.Detail.MethodSignature}");
                }

                foreach (var call in methodStructure.Calls)
                {
                    MethodListEntry calleeEntry;
                    try
                    {
                        calleeEntry = ResolveCallee(index, current, call);
                    }
                    catch (InvalidOperationException ex)
                    {
                        failureDetails.Add(CreateFailureDetail(current, structure.FilePath, call, ex.Message));
                        var failureInsertion = BuildFailureInsertion(structure.OriginalText, call);
                        if (failureInsertion != null)
                        {
                            AddInsertion(modifications, structure.FilePath, failureInsertion);
                            updatedCallCount++;
                        }

                        continue;
                    }

                    if (calleeEntry == null)
                    {
                        continue;
                    }

                    var insertion = BuildInsertion(structure.OriginalText, call, calleeEntry, prefix, normalizedMethodListPath);
                    if (insertion != null)
                    {
                        AddInsertion(modifications, structure.FilePath, insertion);
                        updatedCallCount++;
                    }

                    var calleeKey = CreateMethodKey(calleeEntry);
                    if (!visited.Contains(calleeKey))
                    {
                        stack.Push(calleeEntry);
                    }
                }
            }

            var updatedFiles = new List<string>();
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

            return new TagJumpEmbeddingResult(updatedFiles.Count, updatedCallCount, updatedFiles,
                failureDetails.Count, failureDetails);
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

            if (!map.TryGetValue(filePath, out var list))
            {
                list = new List<TagInsertion>();
                map[filePath] = list;
            }

            list.Add(insertion);
        }

        private static TagInsertion BuildInsertion(string originalText, JavaMethodCallStructure call,
            MethodListEntry calleeEntry, string prefix, string methodListPath)
        {
            if (call == null || calleeEntry == null)
            {
                return null;
            }

            var baseIndex = call.CallEndIndex + 1;
            if (baseIndex < 0 || baseIndex > originalText.Length)
            {
                return null;
            }

            var methodListSegment = methodListPath ?? string.Empty;
            string tagContent;
            if (string.IsNullOrEmpty(methodListSegment))
            {
                tagContent = string.Concat(calleeEntry.Detail.FilePath, "\t", calleeEntry.Detail.MethodSignature);
            }
            else
            {
                tagContent = string.Concat(calleeEntry.Detail.FilePath, "\t", calleeEntry.Detail.MethodSignature, "\t", methodListSegment);
            }
            var replacement = (prefix ?? string.Empty) + tagContent;

            var insertIndex = baseIndex;
            while (insertIndex < originalText.Length && (originalText[insertIndex] == ' ' || originalText[insertIndex] == '\t'))
            {
                insertIndex++;
            }

            if (insertIndex < originalText.Length && originalText[insertIndex] == ';')
            {
                insertIndex++;
            }

            var scanIndex = insertIndex;
            while (scanIndex < originalText.Length)
            {
                var c = originalText[scanIndex];
                if (c == ' ' || c == '\t')
                {
                    scanIndex++;
                    continue;
                }

                break;
            }

            var prefixStart = scanIndex;
            var existingPrefix = TagJumpSyntaxHelper.FindExistingPrefixSegment(originalText, prefixStart, prefix);
            if (existingPrefix.HasValue)
            {
                var length = Math.Min(existingPrefix.Length, originalText.Length - existingPrefix.Start);
                if (length > 0)
                {
                    var existingText = originalText.Substring(existingPrefix.Start, length);
                    if (string.Equals(existingText, replacement, StringComparison.Ordinal))
                    {
                        return null;
                    }
                }

                return new TagInsertion(existingPrefix.Start, existingPrefix.Length, replacement);
            }

            return new TagInsertion(insertIndex, 0, replacement);
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

        private static TagInsertion BuildFailureInsertion(string originalText, JavaMethodCallStructure call)
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
            JavaMethodCallStructure call, string reason)
        {
            var callerSignature = currentMethod?.Detail?.MethodSignature ?? string.Empty;
            var callExpression = BuildCallExpression(call);
            var lineNumber = call?.LineNumber ?? 0;
            return new TagJumpFailureDetail(filePath, lineNumber, callerSignature, callExpression, reason ?? string.Empty);
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

        private static InvalidOperationException BuildAmbiguousCallException(MethodListEntry currentMethod,
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

            return new InvalidOperationException(builder.ToString());
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

                var detail = new MethodDefinitionDetail(filePath, packageName, className, methodSignature);
                var methodName = GetMethodNameFromSignature(methodSignature);
                var parameterCount = CountParametersFromSignature(methodSignature);
                results.Add(new MethodListEntry(detail, filePath, methodName, parameterCount));
            }

            return results;
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
