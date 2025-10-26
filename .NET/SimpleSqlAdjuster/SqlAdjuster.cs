using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SimpleSqlAdjuster
{
    internal sealed class SqlAdjuster
    {
        private readonly SqlFormatter _formatter = new SqlFormatter();

        public string Process(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var outputs = new List<string>();
            using (var reader = new StringReader(input))
            {
                string line;
                var lineNumber = 1;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        outputs.Add(ProcessLine(line, lineNumber));
                    }

                    lineNumber++;
                }
            }

            return string.Join(Environment.NewLine + Environment.NewLine, outputs);
        }

        private string ProcessLine(string rawLine, int lineNumber)
        {
            var trimmed = rawLine.Trim();
            if (string.IsNullOrEmpty(trimmed))
            {
                return string.Empty;
            }

            if (trimmed.StartsWith("--", StringComparison.Ordinal) ||
                trimmed.StartsWith("/*", StringComparison.Ordinal))
            {
                return trimmed;
            }

            string variableName = null;
            string sqlText = trimmed;
            var sqlStartIndex = 0;

            var equalsIndex = trimmed.IndexOf('=');
            if (equalsIndex > 0)
            {
                var left = trimmed.Substring(0, equalsIndex).Trim();
                var right = trimmed.Substring(equalsIndex + 1).Trim();

                if (IsLikelyVariable(left) && LooksLikeSqlStart(right))
                {
                    variableName = left;
                    sqlText = right;
                    sqlStartIndex = FindSqlStartIndex(rawLine, equalsIndex);
                }
            }

            if (!LooksLikeSqlStart(sqlText))
            {
                throw new SqlProcessingException("SQLの開始位置を特定できません。", lineNumber, 1);
            }

            if (variableName == null)
            {
                sqlStartIndex = FindSqlStartIndex(rawLine, -1);
            }

            var columnBase = sqlStartIndex + 1;

            var expanded = MacroExpander.Expand(sqlText, lineNumber, columnBase);
            var formatted = _formatter.Format(expanded, lineNumber);

            if (string.IsNullOrEmpty(variableName))
            {
                return formatted;
            }

            return variableName + Environment.NewLine + formatted;
        }

        private static int FindSqlStartIndex(string rawLine, int equalsIndexInTrimmed)
        {
            if (string.IsNullOrEmpty(rawLine))
            {
                return 0;
            }

            var leadingWhitespaceCount = 0;
            while (leadingWhitespaceCount < rawLine.Length && char.IsWhiteSpace(rawLine[leadingWhitespaceCount]))
            {
                leadingWhitespaceCount++;
            }

            if (equalsIndexInTrimmed >= 0)
            {
                var rawEqualsIndex = rawLine.IndexOf('=', leadingWhitespaceCount);
                if (rawEqualsIndex < 0)
                {
                    rawEqualsIndex = rawLine.IndexOf('=');
                }

                if (rawEqualsIndex >= 0)
                {
                    var sqlStart = rawEqualsIndex + 1;
                    while (sqlStart < rawLine.Length && char.IsWhiteSpace(rawLine[sqlStart]))
                    {
                        sqlStart++;
                    }

                    return Math.Min(sqlStart, rawLine.Length - 1);
                }
            }

            var sqlStartIndex = leadingWhitespaceCount;
            while (sqlStartIndex < rawLine.Length && char.IsWhiteSpace(rawLine[sqlStartIndex]))
            {
                sqlStartIndex++;
            }

            return Math.Min(sqlStartIndex, Math.Max(rawLine.Length - 1, 0));
        }

        private static bool IsLikelyVariable(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (var c in text)
            {
                if (char.IsWhiteSpace(c))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool LooksLikeSqlStart(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            var builder = new StringBuilder();
            foreach (var c in text)
            {
                if (char.IsLetter(c))
                {
                    builder.Append(char.ToUpperInvariant(c));
                }
                else if (builder.Length > 0)
                {
                    break;
                }

                if (builder.Length >= 12)
                {
                    break;
                }
            }

            var keyword = builder.ToString();
            if (keyword.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("WITH", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("UPDATE", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("INSERT", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("DELETE", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("MERGE", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("CREATE", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("DROP", StringComparison.OrdinalIgnoreCase) ||
                keyword.StartsWith("ALTER", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return text.StartsWith("(", StringComparison.Ordinal);
        }
    }

    internal static class MacroExpander
    {
        public static string Expand(string sql, int lineNumber, int columnOffset)
        {
            if (string.IsNullOrEmpty(sql))
            {
                return sql;
            }

            var builder = new StringBuilder(sql.Length);
            var index = 0;
            while (index < sql.Length)
            {
                if (IsWhereMacro(sql, index, lineNumber, columnOffset, out var macro))
                {
                    builder.Append(macro.ExpandedText);
                    index = macro.EndIndex;
                    continue;
                }

                builder.Append(sql[index]);
                index++;
            }

            return builder.ToString();
        }

        private static bool IsWhereMacro(string sql, int startIndex, int lineNumber, int columnOffset, out WhereMacro macro)
        {
            macro = default(WhereMacro);

            if (startIndex + 8 > sql.Length)
            {
                return false;
            }

            if (!string.Equals(sql.Substring(startIndex, 8), ":_WHERE_", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var typeStart = startIndex + 8;
            if (typeStart >= sql.Length)
            {
                throw new SqlProcessingException("WHEREマクロの形式が不正です。", lineNumber, columnOffset + typeStart);
            }

            var typeEnd = typeStart;
            while (typeEnd < sql.Length && char.IsLetter(sql[typeEnd]))
            {
                typeEnd++;
            }

            var macroType = sql.Substring(typeStart, typeEnd - typeStart).ToUpperInvariant();
            if (macroType != "AND" && macroType != "OR")
            {
                throw new SqlProcessingException("WHEREマクロの種類を判別できません。", lineNumber, columnOffset + typeStart);
            }

            var index = typeEnd;
            while (index < sql.Length && char.IsWhiteSpace(sql[index]))
            {
                index++;
            }

            if (index >= sql.Length || sql[index] != '(')
            {
                throw new SqlProcessingException("WHEREマクロの引数が見つかりません。", lineNumber, columnOffset + index);
            }

            index++;
            var depth = 1;
            var contentStart = index;
            while (index < sql.Length && depth > 0)
            {
                var c = sql[index];
                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                }

                index++;
            }

            if (depth != 0)
            {
                throw new SqlProcessingException("WHEREマクロの括弧が閉じていません。", lineNumber, columnOffset + sql.Length);
            }

            var content = sql.Substring(contentStart, index - contentStart - 1);
            var segments = SplitMacroArguments(content);
            if (segments.Count == 0)
            {
                throw new SqlProcessingException("WHEREマクロの引数が空です。", lineNumber, columnOffset + contentStart);
            }

            var joiner = " " + macroType + " ";
            var expanded = "WHERE " + string.Join(joiner, segments);

            macro = new WhereMacro(startIndex, index, expanded);
            return true;
        }

        private static List<string> SplitMacroArguments(string content)
        {
            var results = new List<string>();
            var depth = 0;
            var segmentStart = 0;
            for (var i = 0; i < content.Length; i++)
            {
                var c = content[i];
                if (c == '(')
                {
                    depth++;
                    continue;
                }

                if (c == ')')
                {
                    depth = Math.Max(0, depth - 1);
                    continue;
                }

                if (c == '.' && depth == 0 && (i + 1 >= content.Length || char.IsWhiteSpace(content[i + 1])))
                {
                    var segment = content.Substring(segmentStart, i - segmentStart).Trim();
                    if (!string.IsNullOrEmpty(segment))
                    {
                        results.Add(segment);
                    }

                    segmentStart = i + 1;
                    while (segmentStart < content.Length && char.IsWhiteSpace(content[segmentStart]))
                    {
                        segmentStart++;
                    }
                }
            }

            var lastSegment = content.Substring(segmentStart).Trim();
            if (!string.IsNullOrEmpty(lastSegment))
            {
                results.Add(lastSegment);
            }

            return results;
        }

        private struct WhereMacro
        {
            public WhereMacro(int startIndex, int endIndex, string expanded)
            {
                StartIndex = startIndex;
                EndIndex = endIndex;
                ExpandedText = expanded;
            }

            public int StartIndex { get; }

            public int EndIndex { get; }

            public string ExpandedText { get; }
        }
    }
}
