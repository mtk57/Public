using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace SimpleMethodCallListCreator
{
    public sealed class CollectFilesService
    {
        public CollectFilesResult CollectFiles(string methodListPath, string startSourceFilePath,
            string startMethodText, string collectRootDirectory, string sourceRootDirectory,
            IProgress<CollectFilesProgressInfo> progress = null, CancellationToken cancellationToken = default)
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

            if (string.IsNullOrWhiteSpace(startMethodText))
            {
                throw new ArgumentException("開始メソッドを入力してください。", nameof(startMethodText));
            }

            if (string.IsNullOrWhiteSpace(collectRootDirectory))
            {
                throw new ArgumentException("収集フォルダパスを入力してください。", nameof(collectRootDirectory));
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

            var normalizedStartSourcePath = NormalizePath(startSourceFilePath);
            if (!File.Exists(normalizedStartSourcePath))
            {
                throw new FileNotFoundException("開始ソースファイルが見つかりません。", normalizedStartSourcePath);
            }

            if (!string.Equals(Path.GetExtension(normalizedStartSourcePath), ".java", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Javaファイル（*.java）のみが対象です。");
            }

            var normalizedCollectRoot = NormalizePath(collectRootDirectory);
            if (!Directory.Exists(normalizedCollectRoot))
            {
                Directory.CreateDirectory(normalizedCollectRoot);
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

            var index = new MethodListIndex(entries);
            var startMethod = ResolveStartMethod(index, normalizedStartSourcePath, startMethodText);
            if (startMethod == null)
            {
                throw new InvalidOperationException("開始メソッドがメソッドリストに見つかりません。");
            }

            var traversal = CollectMethodFiles(index, startMethod, normalizedSourceRoot, cancellationToken);
            return CopyFiles(traversal.Files, normalizedCollectRoot, normalizedSourceRoot, traversal.FailureDetails, progress,
                cancellationToken);
        }

        private static CollectFilesResult CopyFiles(IReadOnlyCollection<string> sourceFiles, string collectRootDirectory,
            string sourceRootDirectory, IReadOnlyList<string> failureDetails,
            IProgress<CollectFilesProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var targets = sourceFiles?.ToList() ?? new List<string>();
            if (targets.Count == 0)
            {
                return new CollectFilesResult(Array.Empty<string>(), 0, 0, collectRootDirectory, failureDetails);
            }

            targets.Sort(StringComparer.OrdinalIgnoreCase);

            var copiedCount = 0;
            var skippedCount = 0;
            var totalCount = targets.Count;

            progress?.Report(new CollectFilesProgressInfo(totalCount, 0, 0, string.Empty));

            foreach (var source in targets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var destination = BuildDestinationPath(source, sourceRootDirectory, collectRootDirectory);
                var destinationDirectory = Path.GetDirectoryName(destination);
                if (!string.IsNullOrEmpty(destinationDirectory) && !Directory.Exists(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                if (File.Exists(destination))
                {
                    skippedCount++;
                    progress?.Report(new CollectFilesProgressInfo(totalCount, copiedCount, skippedCount, destination));
                    continue;
                }

                File.Copy(source, destination, overwrite: false);
                copiedCount++;
                progress?.Report(new CollectFilesProgressInfo(totalCount, copiedCount, skippedCount, destination));
            }

            return new CollectFilesResult(targets, copiedCount, skippedCount, collectRootDirectory, failureDetails);
        }

        private static string BuildDestinationPath(string sourcePath, string sourceRootDirectory,
            string collectRootDirectory)
        {
            var normalizedSource = NormalizePath(sourcePath);
            var normalizedRoot = NormalizePath(sourceRootDirectory);
            if (!normalizedSource.StartsWith(normalizedRoot, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"ソースルートフォルダ外のファイルが検出されました。ファイル: {normalizedSource}, ルート: {normalizedRoot}");
            }

            var relative = normalizedSource.Substring(normalizedRoot.Length)
                .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return Path.Combine(collectRootDirectory, relative);
        }

        private static CollectTraversalResult CollectMethodFiles(MethodListIndex index, MethodListEntry startMethod,
            string sourceRootDirectory, CancellationToken cancellationToken)
        {
            var visited = new HashSet<string>(StringComparer.Ordinal);
            var collectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var failureDetails = new List<string>();
            var analysisCache = new Dictionary<string, JavaFileStructure>(StringComparer.OrdinalIgnoreCase);
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

                if (!current.NormalizedFilePath.StartsWith(sourceRootDirectory, StringComparison.OrdinalIgnoreCase))
                {
                    var errorText = $"ソースルートフォルダ外のファイルを検出しました。ファイル: {current.Detail.FilePath}, ルート: {sourceRootDirectory}";
                    throw new InvalidOperationException(errorText);
                }

                collectedFiles.Add(current.NormalizedFilePath);

                var structure = GetOrAnalyzeStructure(current.NormalizedFilePath, analysisCache);
                var methodStructure = structure.FindMethodBySignature(current.Detail.MethodSignature);
                if (methodStructure == null)
                {
                    throw new InvalidOperationException(
                        $"ソース内にメソッドシグネチャが見つかりませんでした。対象: {current.Detail.MethodSignature}, ファイル: {current.Detail.FilePath}");
                }

                foreach (var call in methodStructure.Calls)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    MethodListEntry calleeEntry;
                    try
                    {
                        calleeEntry = ResolveCallee(index, current, call);
                    }
                    catch (MethodAmbiguityException ex)
                    {
                        failureDetails.Add(BuildFailureText(current, structure.FilePath, call, ex.Message,
                            ex.Candidates));
                        continue;
                    }
                    catch (InvalidOperationException ex)
                    {
                        failureDetails.Add(BuildFailureText(current, structure.FilePath, call, ex.Message));
                        continue;
                    }

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

            return new CollectTraversalResult(collectedFiles, failureDetails);
        }

        private static JavaFileStructure GetOrAnalyzeStructure(string filePath,
            IDictionary<string, JavaFileStructure> cache)
        {
            if (cache.TryGetValue(filePath, out var structure))
            {
                return structure;
            }

            var analyzed = JavaMethodCallAnalyzer.ExtractMethodStructures(filePath);
            cache[filePath] = analyzed;
            return analyzed;
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

        private static MethodListEntry ResolveCallee(MethodListIndex index, MethodListEntry current,
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
                    index.FindByFileAndMethodName(current.NormalizedFilePath, methodName));
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
                throw BuildAmbiguousCallException(current, call, candidates);
            }

            candidates = FilterByArgumentCount(
                new List<MethodListEntry>(index.FindByMethodName(methodName)), call.ArgumentCount);

            if (candidates.Count == 0)
            {
                return null;
            }

            if (candidates.Count > 1)
            {
                throw BuildAmbiguousCallException(current, call, candidates);
            }

            return candidates[0];
        }

        private static MethodAmbiguityException BuildAmbiguousCallException(MethodListEntry current,
            JavaMethodCallStructure call, IReadOnlyCollection<MethodListEntry> candidates)
        {
            var builder = new StringBuilder();
            builder.AppendLine("呼び出し先メソッドを一意に特定できませんでした。");
            builder.AppendLine($"ファイル: {current.Detail.FilePath}");
            builder.AppendLine($"呼び出し元: {current.Detail.MethodSignature}");
            builder.AppendLine($"呼び出し行: {call.LineNumber}");
            builder.AppendLine("候補一覧:");
            var candidateDescriptions = new List<string>();
            foreach (var candidate in candidates)
            {
                builder.Append("  - ");
                builder.Append(candidate.Detail.FilePath);
                builder.Append(" : ");
                builder.Append(candidate.Detail.MethodSignature);
                builder.AppendLine();
                candidateDescriptions.Add($"{candidate.Detail.FilePath} : {candidate.Detail.MethodSignature}");
            }

            return new MethodAmbiguityException(builder.ToString(), candidateDescriptions);
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

                var rawPath = parts[0].Trim();
                if (!Path.IsPathRooted(rawPath))
                {
                    throw new InvalidOperationException($"メソッドリストのファイルパスが絶対パスではありません。（{i + 1}行目）");
                }

                var filePath = NormalizePath(rawPath);
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
                    if (rowText.Length > 0 && int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture,
                        out var parsed) && parsed > 0)
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

        private static string BuildFailureText(MethodListEntry current, string filePath, JavaMethodCallStructure call,
            string reason, IReadOnlyList<string> candidates = null)
        {
            var builder = new StringBuilder();
            builder.AppendLine(reason ?? "メソッド特定に失敗しました。");
            builder.AppendLine($"ファイル: {filePath}");
            builder.AppendLine($"呼び出し元: {current.Detail.MethodSignature}");
            builder.AppendLine($"呼び出し行: {call?.LineNumber ?? 0}");
            if (candidates != null && candidates.Count > 0)
            {
                builder.AppendLine("候補一覧:");
                foreach (var candidate in candidates)
                {
                    builder.Append("  - ");
                    builder.AppendLine(candidate);
                }
            }

            return builder.ToString().TrimEnd();
        }

        private static string CreateMethodKey(MethodListEntry entry)
        {
            return $"{entry.NormalizedFilePath}|{entry.Detail.MethodSignature}";
        }

        private sealed class CollectTraversalResult
        {
            public CollectTraversalResult(IReadOnlyCollection<string> files, IReadOnlyList<string> failureDetails)
            {
                Files = files ?? Array.Empty<string>();
                FailureDetails = failureDetails ?? Array.Empty<string>();
            }

            public IReadOnlyCollection<string> Files { get; }

            public IReadOnlyList<string> FailureDetails { get; }
        }

        private sealed class MethodListEntry
        {
            public MethodListEntry(MethodDefinitionDetail detail, string normalizedFilePath,
                string methodName, int parameterCount)
            {
                Detail = detail ?? throw new ArgumentNullException(nameof(detail));
                NormalizedFilePath = normalizedFilePath ?? throw new ArgumentNullException(nameof(normalizedFilePath));
                MethodName = methodName ?? string.Empty;
                ParameterCount = parameterCount;
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

                if (entries == null)
                {
                    return;
                }

                foreach (var entry in entries)
                {
                    AddToDictionary(_byFilePath, entry.NormalizedFilePath, entry);
                    AddToDictionary(_byClassName, entry.Detail.ClassName, entry);
                    AddToDictionary(_byMethodName, entry.MethodName, entry);
                }
            }

            public List<MethodListEntry> FindByFile(string normalizedFilePath)
            {
                if (_byFilePath.TryGetValue(normalizedFilePath, out var list))
                {
                    return new List<MethodListEntry>(list);
                }

                return new List<MethodListEntry>();
            }

            public List<MethodListEntry> FindByFileAndMethodName(string normalizedFilePath, string methodName)
            {
                var results = new List<MethodListEntry>();
                if (!_byFilePath.TryGetValue(normalizedFilePath, out var list))
                {
                    return results;
                }

                foreach (var entry in list)
                {
                    if (string.Equals(entry.MethodName, methodName, StringComparison.Ordinal))
                    {
                        results.Add(entry);
                    }
                }

                return results;
            }

            public List<MethodListEntry> FindByClassAndMethodName(string className, string methodName)
            {
                var results = new List<MethodListEntry>();
                if (!_byClassName.TryGetValue(className, out var list))
                {
                    return results;
                }

                foreach (var entry in list)
                {
                    if (string.Equals(entry.MethodName, methodName, StringComparison.Ordinal))
                    {
                        results.Add(entry);
                    }
                }

                return results;
            }

            public List<MethodListEntry> FindByMethodName(string methodName)
            {
                if (_byMethodName.TryGetValue(methodName, out var list))
                {
                    return new List<MethodListEntry>(list);
                }

                return new List<MethodListEntry>();
            }

            private static void AddToDictionary(IDictionary<string, List<MethodListEntry>> dictionary,
                string key, MethodListEntry entry)
            {
                if (dictionary == null || entry == null)
                {
                    return;
                }

                if (!dictionary.TryGetValue(key, out var list))
                {
                    list = new List<MethodListEntry>();
                    dictionary[key] = list;
                }

                list.Add(entry);
            }
        }

        public sealed class CollectFilesResult
        {
            public CollectFilesResult(IReadOnlyCollection<string> targets, int copiedCount, int skippedCount,
                string collectRootDirectory, IReadOnlyList<string> failureDetails = null)
            {
                Targets = targets ?? Array.Empty<string>();
                CopiedCount = copiedCount;
                SkippedCount = skippedCount;
                CollectRootDirectory = collectRootDirectory ?? string.Empty;
                FailureDetails = failureDetails ?? Array.Empty<string>();
            }

            public IReadOnlyCollection<string> Targets { get; }

            public int CopiedCount { get; }

            public int SkippedCount { get; }

            public string CollectRootDirectory { get; }

            public int TotalCount
            {
                get { return Targets?.Count ?? 0; }
            }

            public int FailureCount
            {
                get { return FailureDetails?.Count ?? 0; }
            }

            public IReadOnlyList<string> FailureDetails { get; }
        }

        public sealed class CollectFilesProgressInfo
        {
            public CollectFilesProgressInfo(int totalFileCount, int copiedCount, int skippedCount,
                string currentDestinationPath)
            {
                TotalFileCount = totalFileCount;
                CopiedCount = copiedCount;
                SkippedCount = skippedCount;
                CurrentDestinationPath = currentDestinationPath ?? string.Empty;
            }

            public int TotalFileCount { get; }

            public int CopiedCount { get; }

            public int SkippedCount { get; }

            public string CurrentDestinationPath { get; }
        }

        public sealed class MethodAmbiguityException : Exception
        {
            public MethodAmbiguityException(string message, IReadOnlyList<string> candidates)
                : base(message)
            {
                Candidates = candidates ?? Array.Empty<string>();
            }

            public IReadOnlyList<string> Candidates { get; }
        }
    }
}
