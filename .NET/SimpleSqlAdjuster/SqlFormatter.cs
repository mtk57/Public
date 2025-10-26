using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleSqlAdjuster
{
    internal sealed class SqlFormatter
    {
        private const int IndentSize = 2;

        private static readonly ClauseStopper[] SelectStops =
        {
            ClauseStopper.Keyword("FROM"),
            ClauseStopper.Keyword("WHERE"),
            ClauseStopper.Multi("GROUP", "BY"),
            ClauseStopper.Keyword("HAVING"),
            ClauseStopper.Multi("ORDER", "BY"),
            ClauseStopper.Multi("UNION", "ALL"),
            ClauseStopper.Keyword("UNION"),
            ClauseStopper.Keyword("INTERSECT"),
            ClauseStopper.Keyword("MINUS"),
            ClauseStopper.Keyword("EXCEPT")
        };

        private static readonly ClauseStopper[] FromStops =
        {
            ClauseStopper.Keyword("WHERE"),
            ClauseStopper.Multi("GROUP", "BY"),
            ClauseStopper.Keyword("HAVING"),
            ClauseStopper.Multi("ORDER", "BY"),
            ClauseStopper.Multi("UNION", "ALL"),
            ClauseStopper.Keyword("UNION"),
            ClauseStopper.Keyword("INTERSECT"),
            ClauseStopper.Keyword("MINUS"),
            ClauseStopper.Keyword("EXCEPT")
        };

        private static readonly ClauseStopper[] WhereStops = FromStops;
        private static readonly ClauseStopper[] GroupStops =
        {
            ClauseStopper.Keyword("HAVING"),
            ClauseStopper.Multi("ORDER", "BY"),
            ClauseStopper.Multi("UNION", "ALL"),
            ClauseStopper.Keyword("UNION"),
            ClauseStopper.Keyword("INTERSECT"),
            ClauseStopper.Keyword("MINUS"),
            ClauseStopper.Keyword("EXCEPT")
        };

        private static readonly ClauseStopper[] OrderStops =
        {
            ClauseStopper.Multi("UNION", "ALL"),
            ClauseStopper.Keyword("UNION"),
            ClauseStopper.Keyword("INTERSECT"),
            ClauseStopper.Keyword("MINUS"),
            ClauseStopper.Keyword("EXCEPT")
        };

        public string Format(string sql, int lineNumber)
        {
            try
            {
                var tokens = SqlTokenizer.Tokenize(sql);
                if (tokens.Count == 0)
                {
                    return string.Empty;
                }

                var writer = new SqlWriter(IndentSize);
                var formatter = new StatementFormatter(tokens, writer);
                formatter.Format();

                return writer.ToString();
            }
            catch (SqlProcessingException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new SqlProcessingException("SQLの整形に失敗しました。", lineNumber, 1, ex);
            }
        }

        private sealed class StatementFormatter
        {
            private readonly List<SqlToken> _tokens;
            private readonly SqlWriter _writer;
            private int _index;

            public StatementFormatter(List<SqlToken> tokens, SqlWriter writer)
            {
                _tokens = tokens;
                _writer = writer;
            }

            public void Format()
            {
                while (_index < _tokens.Count)
                {
                    var token = _tokens[_index];

                    if (token.Kind == TokenKind.Comment)
                    {
                        _writer.WriteLine(0, token.Value);
                        _index++;
                        continue;
                    }

                    if (token.IsSymbol(";"))
                    {
                        _index++;
                        continue;
                    }

                    if (token.IsKeyword("SELECT"))
                    {
                        _writer.WriteLine(0, "SELECT");
                        _index++;
                        FormatCommaSeparatedClause(SelectStops);
                        continue;
                    }

                    if (token.IsKeyword("FROM"))
                    {
                        _writer.WriteLine(0, "FROM");
                        _index++;
                        FormatFromClause();
                        continue;
                    }

                    if (token.IsKeyword("WHERE"))
                    {
                        _writer.WriteLine(0, "WHERE");
                        _index++;
                        FormatWhereClause();
                        continue;
                    }

                    if (token.IsKeyword("GROUP") && PeekKeyword(1, "BY"))
                    {
                        _writer.WriteLine(0, "GROUP BY");
                        _index += 2;
                        FormatCommaSeparatedClause(GroupStops);
                        continue;
                    }

                    if (token.IsKeyword("ORDER") && PeekKeyword(1, "BY"))
                    {
                        _writer.WriteLine(0, "ORDER BY");
                        _index += 2;
                        FormatCommaSeparatedClause(OrderStops);
                        continue;
                    }

                    if (token.IsKeyword("HAVING"))
                    {
                        _writer.WriteLine(0, "HAVING");
                        _index++;
                        FormatWhereClause();
                        continue;
                    }

                    if (token.IsKeyword("UNION"))
                    {
                        var text = "UNION";
                        if (PeekKeyword(1, "ALL"))
                        {
                            text = "UNION ALL";
                            _index += 2;
                        }
                        else
                        {
                            _index++;
                        }

                        _writer.WriteLine(0, text);
                        continue;
                    }

                    WriteGeneralLine();
                }
            }

            private void FormatFromClause()
            {
                var tokens = CollectTokensUntil(FromStops);
                var lines = SplitFromClause(tokens);
                foreach (var line in lines)
                {
                    _writer.WriteLine(1, line);
                }
            }

            private void FormatWhereClause()
            {
                var segments = new List<WhereSegment>();
                var currentTokens = new List<SqlToken>();
                var depth = 0;
                string pendingOperator = null;

                while (_index < _tokens.Count)
                {
                    if (depth == 0 && MatchesStopper(_index, WhereStops))
                    {
                        break;
                    }

                    var token = _tokens[_index];

                    if (token.Kind == TokenKind.Comment)
                    {
                        currentTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (token.IsSymbol("("))
                    {
                        depth++;
                        currentTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (token.IsSymbol(")"))
                    {
                        depth = Math.Max(0, depth - 1);
                        currentTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (depth == 0 && (token.IsKeyword("AND") || token.IsKeyword("OR")))
                    {
                        segments.Add(new WhereSegment(pendingOperator, currentTokens));
                        pendingOperator = token.UpperValue;
                        currentTokens = new List<SqlToken>();
                        _index++;
                        continue;
                    }

                    currentTokens.Add(token);
                    _index++;
                }

                segments.Add(new WhereSegment(pendingOperator, currentTokens));

                foreach (var segment in segments)
                {
                    if (segment.IsEmpty)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(segment.Operator))
                    {
                        _writer.WriteLine(1, segment.Operator);
                    }

                    var text = RenderTokens(segment.Tokens);
                    if (!string.IsNullOrEmpty(text))
                    {
                        _writer.WriteLine(1, text);
                    }
                }
            }

            private void FormatCommaSeparatedClause(ClauseStopper[] stops)
            {
                var itemTokens = new List<SqlToken>();
                var depth = 0;

                while (_index < _tokens.Count)
                {
                    if (depth == 0 && MatchesStopper(_index, stops))
                    {
                        break;
                    }

                    var token = _tokens[_index];

                    if (token.Kind == TokenKind.Comment)
                    {
                        itemTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (token.IsSymbol("("))
                    {
                        depth++;
                        itemTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (token.IsSymbol(")"))
                    {
                        depth = Math.Max(0, depth - 1);
                        itemTokens.Add(token);
                        _index++;
                        continue;
                    }

                    if (depth == 0 && token.IsSymbol(","))
                    {
                        itemTokens.Add(token);
                        var text = RenderTokens(itemTokens);
                        if (!string.IsNullOrEmpty(text))
                        {
                            _writer.WriteLine(1, text);
                        }

                        itemTokens = new List<SqlToken>();
                        _index++;
                        continue;
                    }

                    itemTokens.Add(token);
                    _index++;
                }

                var tail = RenderTokens(itemTokens);
                if (!string.IsNullOrEmpty(tail))
                {
                    _writer.WriteLine(1, tail);
                }
            }

            private void WriteGeneralLine()
            {
                var tokens = new List<SqlToken>();
                while (_index < _tokens.Count)
                {
                    var token = _tokens[_index];

                    if (token.Kind == TokenKind.Comment || token.IsSymbol(";"))
                    {
                        break;
                    }

                    if (token.IsKeyword("SELECT") ||
                        token.IsKeyword("FROM") ||
                        token.IsKeyword("WHERE") ||
                        token.IsKeyword("HAVING") ||
                        (token.IsKeyword("GROUP") && PeekKeyword(1, "BY")) ||
                        (token.IsKeyword("ORDER") && PeekKeyword(1, "BY")) ||
                        token.IsKeyword("UNION"))
                    {
                        break;
                    }

                    tokens.Add(token);
                    _index++;
                }

                var text = RenderTokens(tokens);
                if (!string.IsNullOrEmpty(text))
                {
                    _writer.WriteLine(0, text);
                }
            }

            private List<SqlToken> CollectTokensUntil(ClauseStopper[] stops)
            {
                var result = new List<SqlToken>();
                var depth = 0;

                while (_index < _tokens.Count)
                {
                    if (depth == 0 && MatchesStopper(_index, stops))
                    {
                        break;
                    }

                    var token = _tokens[_index];
                    result.Add(token);

                    if (token.IsSymbol("("))
                    {
                        depth++;
                    }
                    else if (token.IsSymbol(")"))
                    {
                        depth = Math.Max(0, depth - 1);
                    }

                    _index++;
                }

                return result;
            }

            private List<string> SplitFromClause(List<SqlToken> tokens)
            {
                var lines = new List<string>();
                var buffer = new List<SqlToken>();
                var depth = 0;

                for (var i = 0; i < tokens.Count; i++)
                {
                    var token = tokens[i];

                    if (token.IsSymbol("("))
                    {
                        depth++;
                        buffer.Add(token);
                        continue;
                    }

                    if (token.IsSymbol(")"))
                    {
                        depth = Math.Max(0, depth - 1);
                        buffer.Add(token);
                        continue;
                    }

                    if (depth == 0 && IsJoinKeyword(tokens, i, out var joinLength))
                    {
                        if (buffer.Count > 0)
                        {
                            var current = RenderTokens(buffer);
                            if (!string.IsNullOrEmpty(current))
                            {
                                lines.Add(current);
                            }

                            buffer = new List<SqlToken>();
                        }

                        for (var j = 0; j < joinLength; j++)
                        {
                            buffer.Add(tokens[i + j]);
                        }

                        i += joinLength - 1;
                        continue;
                    }

                    if (depth == 0 && token.IsSymbol(","))
                    {
                        buffer.Add(token);
                        var text = RenderTokens(buffer);
                        if (!string.IsNullOrEmpty(text))
                        {
                            lines.Add(text);
                        }

                        buffer = new List<SqlToken>();
                        continue;
                    }

                    buffer.Add(token);
                }

                var remaining = RenderTokens(buffer);
                if (!string.IsNullOrEmpty(remaining))
                {
                    lines.Add(remaining);
                }

                return lines;
            }

            private bool MatchesStopper(int index, ClauseStopper[] stops)
            {
                if (index >= _tokens.Count)
                {
                    return true;
                }

                if (_tokens[index].IsSymbol(";"))
                {
                    return true;
                }

                foreach (var stopper in stops)
                {
                    if (stopper != null && stopper.IsMatch(_tokens, index))
                    {
                        return true;
                    }
                }

                return false;
            }

            private bool PeekKeyword(int offset, string keyword)
            {
                var target = _index + offset;
                if (target >= _tokens.Count)
                {
                    return false;
                }

                return _tokens[target].IsKeyword(keyword);
            }

            private bool IsJoinKeyword(List<SqlToken> tokens, int index, out int length)
            {
                length = 0;
                if (index >= tokens.Count)
                {
                    return false;
                }

                if (tokens[index].IsKeyword("JOIN"))
                {
                    length = 1;
                    return true;
                }

                if (tokens[index].IsKeyword("INNER") && index + 1 < tokens.Count && tokens[index + 1].IsKeyword("JOIN"))
                {
                    length = 2;
                    return true;
                }

                if ((tokens[index].IsKeyword("LEFT") || tokens[index].IsKeyword("RIGHT") || tokens[index].IsKeyword("FULL")) &&
                    index + 1 < tokens.Count)
                {
                    var len = 1;
                    if (tokens[index + len].IsKeyword("OUTER"))
                    {
                        len++;
                    }

                    if (index + len < tokens.Count && tokens[index + len].IsKeyword("JOIN"))
                    {
                        length = len + 1;
                        return true;
                    }
                }

                if (tokens[index].IsKeyword("CROSS") && index + 1 < tokens.Count && tokens[index + 1].IsKeyword("JOIN"))
                {
                    length = 2;
                    return true;
                }

                return false;
            }
        }

        private static string RenderTokens(IList<SqlToken> tokens)
        {
            if (tokens == null || tokens.Count == 0)
            {
                return string.Empty;
            }

            var builder = new StringBuilder();
            SqlToken previous = null;
            foreach (var token in tokens)
            {
                if (token == null)
                {
                    continue;
                }

                var text = token.Kind == TokenKind.Keyword ? token.UpperValue : token.Value;
                if (builder.Length > 0 && NeedsSpace(previous, token))
                {
                    builder.Append(' ');
                }

                builder.Append(text);
                previous = token;
            }

            return builder.ToString().Trim();
        }

        private static bool NeedsSpace(SqlToken previous, SqlToken current)
        {
            if (previous == null || current == null)
            {
                return false;
            }

            if (previous.Kind == TokenKind.Comment || current.Kind == TokenKind.Comment)
            {
                return true;
            }

            if (current.IsSymbol(")") || current.IsSymbol(",") || current.IsSymbol(".") || current.IsSymbol(";"))
            {
                return false;
            }

            if (current.IsSymbol("("))
            {
                return false;
            }

            if (previous.IsSymbol("(") || previous.IsSymbol("."))
            {
                return false;
            }

            if (previous.IsSymbol(","))
            {
                return true;
            }

            return true;
        }

        private sealed class WhereSegment
        {
            public WhereSegment(string op, List<SqlToken> tokens)
            {
                Operator = op;
                Tokens = tokens ?? new List<SqlToken>();
            }

            public string Operator { get; }

            public List<SqlToken> Tokens { get; }

            public bool IsEmpty
            {
                get
                {
                    return Tokens == null || Tokens.Count == 0;
                }
            }
        }
    }

    internal sealed class SqlWriter
    {
        private readonly StringBuilder _builder = new StringBuilder();
        private readonly int _indentSize;

        public SqlWriter(int indentSize)
        {
            _indentSize = Math.Max(0, indentSize);
        }

        public void WriteLine(int indentLevel, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var lines = text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                _builder.Append(new string(' ', Math.Max(0, indentLevel) * _indentSize));
                _builder.AppendLine(line.TrimEnd());
            }
        }

        public override string ToString()
        {
            return _builder.ToString().TrimEnd();
        }
    }

    internal sealed class ClauseStopper
    {
        private readonly string[] _keywords;

        private ClauseStopper(params string[] keywords)
        {
            _keywords = keywords;
        }

        public static ClauseStopper Keyword(string keyword)
        {
            return new ClauseStopper(keyword);
        }

        public static ClauseStopper Multi(params string[] keywords)
        {
            return new ClauseStopper(keywords);
        }

        public bool IsMatch(List<SqlToken> tokens, int index)
        {
            if (_keywords == null || _keywords.Length == 0)
            {
                return false;
            }

            for (var i = 0; i < _keywords.Length; i++)
            {
                var position = index + i;
                if (position >= tokens.Count)
                {
                    return false;
                }

                if (!tokens[position].IsKeyword(_keywords[i]))
                {
                    return false;
                }
            }

            return true;
        }
    }

    internal enum TokenKind
    {
        Keyword,
        Identifier,
        Symbol,
        StringLiteral,
        Number,
        Parameter,
        Comment
    }

    internal sealed class SqlToken
    {
        public SqlToken(string value, TokenKind kind)
        {
            Value = value;
            Kind = kind;
            UpperValue = value != null ? value.ToUpperInvariant() : string.Empty;
        }

        public string Value { get; }

        public string UpperValue { get; }

        public TokenKind Kind { get; }

        public bool IsKeyword(string keyword)
        {
            return Kind == TokenKind.Keyword && string.Equals(UpperValue, keyword, StringComparison.Ordinal);
        }

        public bool IsSymbol(string symbol)
        {
            return Kind == TokenKind.Symbol && string.Equals(Value, symbol, StringComparison.Ordinal);
        }
    }

    internal static class SqlTokenizer
    {
        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "SELECT",
            "FROM",
            "WHERE",
            "GROUP",
            "BY",
            "ORDER",
            "HAVING",
            "UNION",
            "ALL",
            "DISTINCT",
            "AND",
            "OR",
            "NOT",
            "INNER",
            "LEFT",
            "RIGHT",
            "FULL",
            "OUTER",
            "JOIN",
            "ON",
            "INSERT",
            "UPDATE",
            "DELETE",
            "INTO",
            "VALUES",
            "SET",
            "WITH",
            "AS",
            "CASE",
            "WHEN",
            "THEN",
            "ELSE",
            "END",
            "OVER",
            "PARTITION",
            "CROSS",
            "APPLY",
            "MINUS",
            "INTERSECT",
            "EXCEPT",
            "USING",
            "MERGE"
        };

        public static List<SqlToken> Tokenize(string sql)
        {
            var tokens = new List<SqlToken>();
            if (string.IsNullOrEmpty(sql))
            {
                return tokens;
            }

            var length = sql.Length;
            var index = 0;
            while (index < length)
            {
                var ch = sql[index];

                if (char.IsWhiteSpace(ch))
                {
                    index++;
                    continue;
                }

                if (ch == '-' && index + 1 < length && sql[index + 1] == '-')
                {
                    var start = index;
                    index += 2;
                    while (index < length && sql[index] != '\n' && sql[index] != '\r')
                    {
                        index++;
                    }

                    tokens.Add(new SqlToken(sql.Substring(start, index - start), TokenKind.Comment));
                    continue;
                }

                if (ch == '/' && index + 1 < length && sql[index + 1] == '*')
                {
                    var start = index;
                    index += 2;
                    while (index + 1 < length && !(sql[index] == '*' && sql[index + 1] == '/'))
                    {
                        index++;
                    }

                    index = Math.Min(length, index + 2);
                    tokens.Add(new SqlToken(sql.Substring(start, Math.Min(length, index) - start), TokenKind.Comment));
                    continue;
                }

                if (ch == '\'' || ch == '"')
                {
                    var quote = ch;
                    var start = index;
                    index++;
                    while (index < length)
                    {
                        if (sql[index] == quote)
                        {
                            index++;
                            if (index < length && sql[index] == quote)
                            {
                                index++;
                                continue;
                            }

                            break;
                        }

                        index++;
                    }

                    tokens.Add(new SqlToken(sql.Substring(start, Math.Min(length, index) - start), TokenKind.StringLiteral));
                    continue;
                }

                if (ch == ':' || ch == '@')
                {
                    var start = index;
                    index++;
                    while (index < length)
                    {
                        var c = sql[index];
                        if (char.IsLetterOrDigit(c) || c == '_' || c == '$' || c == '#')
                        {
                            index++;
                            continue;
                        }

                        break;
                    }

                    tokens.Add(new SqlToken(sql.Substring(start, index - start), TokenKind.Parameter));
                    continue;
                }

                if (char.IsLetter(ch) || ch == '_' || ch == '$' || ch == '#')
                {
                    var start = index;
                    index++;
                    while (index < length)
                    {
                        var c = sql[index];
                        if (char.IsLetterOrDigit(c) || c == '_' || c == '$' || c == '#')
                        {
                            index++;
                            continue;
                        }

                        break;
                    }

                    var word = sql.Substring(start, index - start);
                    tokens.Add(new SqlToken(word, Keywords.Contains(word) ? TokenKind.Keyword : TokenKind.Identifier));
                    continue;
                }

                if (char.IsDigit(ch))
                {
                    var start = index;
                    index++;
                    while (index < length && (char.IsDigit(sql[index]) || sql[index] == '.'))
                    {
                        index++;
                    }

                    tokens.Add(new SqlToken(sql.Substring(start, index - start), TokenKind.Number));
                    continue;
                }

                if (index + 1 < length)
                {
                    var two = sql.Substring(index, 2);
                    if (two == "<=" || two == ">=" || two == "<>" || two == "!=")
                    {
                        tokens.Add(new SqlToken(two, TokenKind.Symbol));
                        index += 2;
                        continue;
                    }
                }

                tokens.Add(new SqlToken(sql[index].ToString(), TokenKind.Symbol));
                index++;
            }

            return tokens;
        }
    }
}
