using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SimpleSqlAdjuster
{
    internal sealed class SqlAdjuster
    {
        private readonly SqlFormatter _formatter = new SqlFormatter();

        public string Process(string input, bool appendTableMetadata = false)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var outputs = new List<string>();
            var formatterOptions = appendTableMetadata
                ? new SqlFormatterOptions { AppendTableMetadata = true }
                : null;
            using (var reader = new StringReader(input))
            {
                string line;
                var lineNumber = 1;
                while ((line = reader.ReadLine()) != null)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        outputs.Add(ProcessLine(line, lineNumber, formatterOptions));
                    }

                    lineNumber++;
                }
            }

            return string.Join(Environment.NewLine + Environment.NewLine, outputs);
        }

        private string ProcessLine(string rawLine, int lineNumber, SqlFormatterOptions formatterOptions)
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

            var expansion = MacroExpander.Expand(sqlText, lineNumber, columnBase);
            var formatted = _formatter.Format(expansion.Sql, lineNumber, formatterOptions);
            formatted = MacroExpander.ApplyMacroFormatting(formatted, expansion.Macros);

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
        public static MacroExpansionResult Expand(string sql, int lineNumber, int columnOffset)
        {
            if (string.IsNullOrEmpty(sql))
            {
                return new MacroExpansionResult(sql ?? string.Empty, new List<WhereMacroReplacement>());
            }

            var builder = new StringBuilder(sql.Length);
            var macros = new List<WhereMacroReplacement>();
            var index = 0;
            var placeholderIndex = 0;

            while (index < sql.Length)
            {
                if (TryParseWhereMacro(sql, index, lineNumber, columnOffset, out var macro))
                {
                    var placeholder = "__SIMPLE_SQL_ADJUSTER_MACRO_" + placeholderIndex + "__";
                    placeholderIndex++;

                    if (builder.Length > 0 && !char.IsWhiteSpace(builder[builder.Length - 1]))
                    {
                        builder.Append(' ');
                    }

                    builder.Append("WHERE ");
                    builder.Append(placeholder);

                    macros.Add(new WhereMacroReplacement(placeholder, macro.MacroType, macro.Segments));
                    index = macro.EndIndex;
                    continue;
                }

                builder.Append(sql[index]);
                index++;
            }

            return new MacroExpansionResult(builder.ToString(), macros);
        }

        public static string ApplyMacroFormatting(string formattedSql, IList<WhereMacroReplacement> macros)
        {
            if (string.IsNullOrEmpty(formattedSql) || macros == null || macros.Count == 0)
            {
                return formattedSql;
            }

            var lookup = new Dictionary<string, WhereMacroReplacement>(StringComparer.Ordinal);
            foreach (var macro in macros)
            {
                lookup[macro.Placeholder] = macro;
            }

            var lines = formattedSql.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var resultLines = new List<string>(lines.Length);

            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.Equals(line.Trim(), "WHERE", StringComparison.OrdinalIgnoreCase) && i + 1 < lines.Length)
                {
                    var nextLine = lines[i + 1];
                    var placeholderKey = nextLine.Trim();
                    if (lookup.TryGetValue(placeholderKey, out var replacement))
                    {
                        AppendMacroLines(resultLines, replacement);
                        i++; // skip placeholder line
                        continue;
                    }
                }

                resultLines.Add(line);
            }

            return string.Join(Environment.NewLine, resultLines).TrimEnd();
        }

        private static void AppendMacroLines(List<string> lines, WhereMacroReplacement replacement)
        {
            lines.Add(":_WHERE_" + replacement.MacroType);
            lines.Add("(");

            var count = replacement.Segments.Count;
            for (var i = 0; i < count; i++)
            {
                var segment = replacement.Segments[i];
                var suffix = i < count - 1 ? "," : string.Empty;
                lines.Add("  " + segment + suffix);
            }

            lines.Add(")");
        }

        private static bool TryParseWhereMacro(string sql, int startIndex, int lineNumber, int columnOffset, out WhereMacroParseResult macro)
        {
            macro = default(WhereMacroParseResult);

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

            macro = new WhereMacroParseResult(index, macroType, segments);
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

                var isDotDelimiter = c == '.' && (i + 1 >= content.Length || char.IsWhiteSpace(content[i + 1]));
                var isCommaDelimiter = c == ','; // カンマは常に区切りとする

                if (depth == 0 && (isDotDelimiter || isCommaDelimiter))
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

        private struct WhereMacroParseResult
        {
            public WhereMacroParseResult(int endIndex, string macroType, List<string> segments)
            {
                EndIndex = endIndex;
                MacroType = macroType;
                Segments = segments ?? new List<string>();
            }

            public int EndIndex { get; }

            public string MacroType { get; }

            public List<string> Segments { get; }
        }

        public sealed class MacroExpansionResult
        {
            public MacroExpansionResult(string sql, List<WhereMacroReplacement> macros)
            {
                Sql = sql ?? string.Empty;
                Macros = macros ?? new List<WhereMacroReplacement>();
            }

            public string Sql { get; }

            public IList<WhereMacroReplacement> Macros { get; }
        }

        public sealed class WhereMacroReplacement
        {
            public WhereMacroReplacement(string placeholder, string macroType, List<string> segments)
            {
                Placeholder = placeholder;
                MacroType = macroType;
                Segments = segments ?? new List<string>();
            }

            public string Placeholder { get; }

            public string MacroType { get; }

            public List<string> Segments { get; }
        }
    }
}
