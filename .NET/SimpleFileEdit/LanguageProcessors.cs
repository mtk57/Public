using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleFileSearch
{
    public sealed class MethodSpanInfo
    {
        public int StartLine { get; set; }
        public int EndLine { get; set; }
        public string RawSignature { get; set; }
    }

    public interface ILanguageProcessor
    {
        List<MethodSpanInfo> FindMethodSpans(string[] lines);
        bool IsImportLine(string line);
    }

    public static class LanguageProcessorFactory
    {
        public static ILanguageProcessor Create(string extension)
        {
            if (string.Equals(extension, ".java", StringComparison.OrdinalIgnoreCase))
            {
                return new JavaLanguageProcessor();
            }

            return null;
        }
    }

    public sealed class JavaLanguageProcessor : ILanguageProcessor
    {
        public bool IsImportLine(string line)
        {
            string trimmed = (line ?? string.Empty).TrimStart();
            return trimmed.StartsWith("import ", StringComparison.Ordinal);
        }

        public List<MethodSpanInfo> FindMethodSpans(string[] lines)
        {
            var spans = new List<MethodSpanInfo>();
            var stack = new Stack<MethodBuildInfo>();
            bool inBlockComment = false;
            bool capturing = false;
            int sigStartLine = -1;
            var sigBuilder = new StringBuilder();
            var rawSigBuilder = new StringBuilder();
            int braceDepth = 0;
            int annotStartLine = -1;

            for (int i = 0; i < lines.Length; i++)
            {
                string rawLine = lines[i] ?? string.Empty;
                string stripped = StripLineComments(rawLine, ref inBlockComment);
                string braceSafe = StripStringLiterals(stripped);
                string trimmed = stripped.Trim();

                if (!capturing && IsAnnotationLine(trimmed))
                {
                    if (annotStartLine < 0)
                    {
                        annotStartLine = i;
                    }
                }
                else
                {
                    if (!capturing && IsMethodCandidate(trimmed))
                    {
                        capturing = true;
                        sigBuilder.Clear();
                        sigBuilder.Append(trimmed);
                        rawSigBuilder.Clear();
                        rawSigBuilder.Append(rawLine);
                        sigStartLine = i;
                        if (annotStartLine < 0)
                        {
                            annotStartLine = i;
                        }
                    }
                    else if (capturing && !string.IsNullOrWhiteSpace(trimmed))
                    {
                        sigBuilder.Append(" ").Append(trimmed);
                        rawSigBuilder.AppendLine().Append(rawLine);
                    }
                    else if (!capturing)
                    {
                        if (!string.IsNullOrWhiteSpace(trimmed))
                        {
                            annotStartLine = -1;
                        }
                    }

                    if (capturing)
                    {
                        string candidate = sigBuilder.ToString();
                        if (candidate.Contains(";") && !candidate.Contains("{"))
                        {
                            capturing = false;
                            annotStartLine = -1;
                        }
                        else if (candidate.Contains("{"))
                        {
                            stack.Push(new MethodBuildInfo
                            {
                                StartLine = annotStartLine >= 0 ? annotStartLine : sigStartLine,
                                StartDepth = braceDepth,
                                RawSignature = rawSigBuilder.ToString()
                            });
                            capturing = false;
                            annotStartLine = -1;
                        }
                    }
                }

                braceDepth += CountOccurrences(braceSafe, '{');
                braceDepth -= CountOccurrences(braceSafe, '}');
                if (braceDepth < 0)
                {
                    braceDepth = 0;
                }

                while (stack.Count > 0 && braceDepth <= stack.Peek().StartDepth)
                {
                    var b = stack.Pop();
                    spans.Add(new MethodSpanInfo
                    {
                        StartLine = b.StartLine,
                        EndLine = i,
                        RawSignature = b.RawSignature
                    });
                }
            }

            while (stack.Count > 0)
            {
                var b = stack.Pop();
                spans.Add(new MethodSpanInfo
                {
                    StartLine = b.StartLine,
                    EndLine = lines.Length - 1,
                    RawSignature = b.RawSignature
                });
            }

            return spans;
        }

        private static bool IsAnnotationLine(string trimmed)
        {
            return !string.IsNullOrWhiteSpace(trimmed) && trimmed[0] == '@';
        }

        private static bool IsMethodCandidate(string trimmed)
        {
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return false;
            }

            if (trimmed.StartsWith("@", StringComparison.Ordinal))
            {
                return false;
            }

            int parenIndex = trimmed.IndexOf('(');
            if (parenIndex < 0)
            {
                return false;
            }

            string before = trimmed.Substring(0, parenIndex);
            if (before.IndexOf(" class ", StringComparison.Ordinal) >= 0 ||
                before.IndexOf(" interface ", StringComparison.Ordinal) >= 0 ||
                before.IndexOf(" enum ", StringComparison.Ordinal) >= 0)
            {
                return false;
            }

            if (before.EndsWith("class", StringComparison.Ordinal) ||
                before.EndsWith("interface", StringComparison.Ordinal) ||
                before.EndsWith("enum", StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }

        private static string StripLineComments(string line, ref bool inBlockComment)
        {
            if (string.IsNullOrEmpty(line))
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            bool inString = false;
            char stringDelimiter = '\0';

            for (int i = 0; i < line.Length; i++)
            {
                char current = line[i];

                if (inBlockComment)
                {
                    if (current == '*' && i + 1 < line.Length && line[i + 1] == '/')
                    {
                        inBlockComment = false;
                        i++;
                    }
                    continue;
                }

                if (inString)
                {
                    sb.Append(current);
                    if (current == stringDelimiter && (i == 0 || line[i - 1] != '\\'))
                    {
                        inString = false;
                        stringDelimiter = '\0';
                    }
                    continue;
                }

                if (current == '"' || current == '\'')
                {
                    inString = true;
                    stringDelimiter = current;
                    sb.Append(current);
                    continue;
                }

                if (current == '/' && i + 1 < line.Length)
                {
                    char next = line[i + 1];
                    if (next == '/')
                    {
                        break;
                    }
                    if (next == '*')
                    {
                        inBlockComment = true;
                        i++;
                        continue;
                    }
                }

                sb.Append(current);
            }

            return sb.ToString();
        }

        private static string StripStringLiterals(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            bool inString = false;
            char delimiter = '\0';

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (!inString && (c == '"' || c == '\''))
                {
                    inString = true;
                    delimiter = c;
                    sb.Append(' ');
                    continue;
                }

                if (inString)
                {
                    if (c == delimiter && (i == 0 || text[i - 1] != '\\'))
                    {
                        inString = false;
                        delimiter = '\0';
                    }
                    sb.Append(' ');
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private static int CountOccurrences(string text, char target)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int count = 0;
            foreach (char c in text)
            {
                if (c == target)
                {
                    count++;
                }
            }

            return count;
        }

        private sealed class MethodBuildInfo
        {
            public int StartLine { get; set; }
            public int StartDepth { get; set; }
            public string RawSignature { get; set; }
        }
    }
}
