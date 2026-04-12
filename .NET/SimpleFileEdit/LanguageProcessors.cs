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
            bool inBlockComment = false;
            int depth = 0;
            int blockStartLine = -1;
            string blockRawSignature = null;

            for (int i = 0; i < lines.Length; i++)
            {
                string rawLine = lines[i] ?? string.Empty;
                string stripped = StripLineComments(rawLine, ref inBlockComment);
                string safe = StripStringLiterals(stripped);

                int opens = CountOccurrences(safe, '{');
                int closes = CountOccurrences(safe, '}');
                int prevDepth = depth;

                if (prevDepth == 1 && opens > 0)
                {
                    int bracePos = safe.IndexOf('{');
                    string beforeBrace = bracePos >= 0 ? safe.Substring(0, bracePos) : string.Empty;

                    if (!beforeBrace.Contains("="))
                    {
                        int declStart = FindDeclarationStart(lines, i);
                        var sigBuilder = new StringBuilder();
                        for (int k = declStart; k <= i; k++)
                        {
                            if (k > declStart)
                            {
                                sigBuilder.AppendLine();
                            }
                            sigBuilder.Append(lines[k]);
                        }
                        blockStartLine = declStart;
                        blockRawSignature = sigBuilder.ToString();
                    }
                }

                depth += opens - closes;
                if (depth < 0)
                {
                    depth = 0;
                }

                if (depth == 1 && blockStartLine >= 0 && closes > 0)
                {
                    spans.Add(new MethodSpanInfo
                    {
                        StartLine = blockStartLine,
                        EndLine = i,
                        RawSignature = blockRawSignature
                    });
                    blockStartLine = -1;
                    blockRawSignature = null;
                }
            }

            if (blockStartLine >= 0)
            {
                spans.Add(new MethodSpanInfo
                {
                    StartLine = blockStartLine,
                    EndLine = lines.Length - 1,
                    RawSignature = blockRawSignature
                });
            }

            return spans;
        }

        private static int FindDeclarationStart(string[] lines, int braceLineIndex)
        {
            int start = braceLineIndex;

            // Phase 1: シグネチャ・アノテーション・直接隣接するコメントを含める
            for (int i = braceLineIndex - 1; i >= 0; i--)
            {
                string trimmed = lines[i].Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    break;
                }
                if (trimmed.EndsWith("}", StringComparison.Ordinal))
                {
                    break;
                }
                if (trimmed.EndsWith(";", StringComparison.Ordinal))
                {
                    break;
                }
                start = i;
            }

            // Phase 2: 空行を挟んだブロックコメント（Javadoc等）を含める
            int searchIndex = start - 1;
            while (searchIndex >= 0 && string.IsNullOrWhiteSpace(lines[searchIndex].Trim()))
            {
                searchIndex--;
            }

            if (searchIndex >= 0 && lines[searchIndex].Trim().EndsWith("*/", StringComparison.Ordinal))
            {
                for (int i = searchIndex; i >= 0; i--)
                {
                    start = i;
                    if (lines[i].Trim().StartsWith("/*", StringComparison.Ordinal))
                    {
                        break;
                    }
                }
            }

            return start;
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
    }
}
