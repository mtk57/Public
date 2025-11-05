using System;
using System.Collections.Generic;

namespace SimpleMethodCallListCreator
{
    internal static class TagJumpSyntaxHelper
    {
        private static readonly string[] DefaultPrefixes = { "//@ ", "//@" };

        public static PrefixSegment FindExistingPrefixSegment(string text, int position, string configuredPrefix)
        {
            if (string.IsNullOrEmpty(text) || position < 0 || position >= text.Length)
            {
                return PrefixSegment.Empty;
            }

            foreach (var candidate in EnumeratePrefixCandidates(configuredPrefix))
            {
                if (StartsWithAt(text, position, candidate))
                {
                    var end = position + candidate.Length;
                    while (end < text.Length)
                    {
                        var ch = text[end];
                        if (ch == '\r' || ch == '\n')
                        {
                            break;
                        }

                        end++;
                    }

                    return new PrefixSegment(position, end - position);
                }
            }

            return PrefixSegment.Empty;
        }

        public static PrefixSegment FindFirstPrefixSegmentInRange(string text, int startIndex, int endIndex, string configuredPrefix)
        {
            if (string.IsNullOrEmpty(text))
            {
                return PrefixSegment.Empty;
            }

            if (startIndex < 0)
            {
                startIndex = 0;
            }

            if (endIndex > text.Length)
            {
                endIndex = text.Length;
            }

            if (startIndex >= endIndex)
            {
                return PrefixSegment.Empty;
            }

            var candidates = GetPrefixCandidates(configuredPrefix);
            if (candidates.Count == 0)
            {
                return PrefixSegment.Empty;
            }

            for (var index = startIndex; index < endIndex; index++)
            {
                foreach (var candidate in candidates)
                {
                    if (!StartsWithAt(text, index, candidate))
                    {
                        continue;
                    }

                    var end = index + candidate.Length;
                    while (end < text.Length)
                    {
                        var ch = text[end];
                        if (ch == '\r' || ch == '\n')
                        {
                            break;
                        }

                        end++;
                    }

                    if (end > endIndex)
                    {
                        end = endIndex;
                    }

                    return new PrefixSegment(index, end - index);
                }
            }

            return PrefixSegment.Empty;
        }

        public static bool TryParseTagLine(string lineText, string configuredPrefix,
            out string filePath, out string methodSignature, out string methodListPath)
        {
            filePath = string.Empty;
            methodSignature = string.Empty;
            methodListPath = string.Empty;

            if (string.IsNullOrEmpty(lineText))
            {
                return false;
            }

            var prefixIndex = -1;
            string matchedPrefix = null;

            foreach (var candidate in EnumeratePrefixCandidates(configuredPrefix))
            {
                var index = lineText.LastIndexOf(candidate, StringComparison.Ordinal);
                if (index >= 0)
                {
                    prefixIndex = index;
                    matchedPrefix = candidate;
                    break;
                }
            }

            if (prefixIndex < 0 || string.IsNullOrEmpty(matchedPrefix))
            {
                return false;
            }

            var contentStart = prefixIndex + matchedPrefix.Length;
            while (contentStart < lineText.Length && lineText[contentStart] == ' ')
            {
                contentStart++;
            }

            if (contentStart >= lineText.Length)
            {
                return false;
            }

            var content = lineText.Substring(contentStart).Trim();
            if (content.Length == 0)
            {
                return false;
            }

            var parts = content.Split('\t');
            if (parts.Length < 2)
            {
                return false;
            }

            var pathPart = parts[0].Trim();
            var signaturePart = parts[1].Trim();
            var methodListPart = string.Empty;
            if (parts.Length >= 3)
            {
                methodListPart = parts[2].Trim();
            }

            if (pathPart.Length == 0 || signaturePart.Length == 0)
            {
                return false;
            }

            filePath = pathPart;
            methodSignature = signaturePart;
            methodListPath = methodListPart;
            return true;
        }

        private static List<string> GetPrefixCandidates(string configuredPrefix)
        {
            var list = new List<string>();
            foreach (var candidate in EnumeratePrefixCandidates(configuredPrefix))
            {
                if (!string.IsNullOrEmpty(candidate))
                {
                    list.Add(candidate);
                }
            }

            return list;
        }

        private static IEnumerable<string> EnumeratePrefixCandidates(string configuredPrefix)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);

            if (!string.IsNullOrEmpty(configuredPrefix))
            {
                if (seen.Add(configuredPrefix))
                {
                    yield return configuredPrefix;
                }

                var trimmed = configuredPrefix.TrimEnd();
                if (!string.Equals(trimmed, configuredPrefix, StringComparison.Ordinal) && trimmed.Length > 0)
                {
                    if (seen.Add(trimmed))
                    {
                        yield return trimmed;
                    }
                }
            }

            foreach (var candidate in DefaultPrefixes)
            {
                if (seen.Add(candidate))
                {
                    yield return candidate;
                }
            }
        }

        private static bool StartsWithAt(string text, int position, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            if (position < 0 || position + value.Length > text.Length)
            {
                return false;
            }

            return string.Compare(text, position, value, 0, value.Length, StringComparison.Ordinal) == 0;
        }

        internal readonly struct PrefixSegment
        {
            public static readonly PrefixSegment Empty = new PrefixSegment(-1, 0);

            public PrefixSegment(int start, int length)
            {
                Start = start;
                Length = length < 0 ? 0 : length;
            }

            public int Start { get; }
            public int Length { get; }

            public bool HasValue => Start >= 0 && Length > 0;
        }
    }
}
