using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleGrep.Core
{
    internal interface IMethodSignatureResolver
    {
        string GetMethodSignature(int lineNumber);
    }

    internal sealed class JavaMethodSignatureResolver : IMethodSignatureResolver
    {
        private static readonly HashSet<string> MethodModifiers = new HashSet<string>(StringComparer.Ordinal)
        {
            "public","protected","private","static","final","abstract","synchronized","native","strictfp","default"
        };

        private static readonly HashSet<string> DisallowedMethodNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "if","for","while","switch","catch","try","return","new","else","do"
        };

        private readonly List<MethodSpan> _methodSpans;

        public JavaMethodSignatureResolver(string[] lines)
        {
            _methodSpans = BuildSpans(lines);
        }

        public string GetMethodSignature(int lineNumber)
        {
            if (_methodSpans.Count == 0)
            {
                return string.Empty;
            }

            int low = 0;
            int high = _methodSpans.Count - 1;
            while (low <= high)
            {
                int mid = (low + high) / 2;
                var span = _methodSpans[mid];
                if (lineNumber < span.StartLine)
                {
                    high = mid - 1;
                }
                else if (lineNumber > span.EndLine)
                {
                    low = mid + 1;
                }
                else
                {
                    return span.Signature;
                }
            }

            return string.Empty;
        }

        private static List<MethodSpan> BuildSpans(string[] lines)
        {
            var spans = new List<MethodSpan>();
            var methodStack = new Stack<MethodSpanBuilder>();
            bool inBlockComment = false;
            bool capturingSignature = false;
            int signatureStartLine = -1;
            var signatureBuilder = new StringBuilder();
            int braceDepth = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                string rawLine = lines[i] ?? string.Empty;
                string withoutComments = RemoveComments(rawLine, ref inBlockComment);
                string braceSafeLine = RemoveStringLiterals(withoutComments);
                string trimmed = withoutComments.Trim();

                if (!capturingSignature && StartsWithAnnotation(trimmed))
                {
                    // アノテーションは次行の署名判定に影響させない
                }
                else
                {
                    if (!capturingSignature && ContainsMethodToken(trimmed))
                    {
                        capturingSignature = true;
                        signatureBuilder.Clear();
                        signatureBuilder.Append(trimmed);
                        signatureStartLine = i + 1;
                    }
                    else if (capturingSignature && !string.IsNullOrWhiteSpace(trimmed))
                    {
                        signatureBuilder.Append(" ").Append(trimmed);
                    }

                    if (capturingSignature)
                    {
                        string candidate = signatureBuilder.ToString();
                        if (candidate.Contains(";") && !candidate.Contains("{"))
                        {
                            // インターフェース等の宣言は対象外
                            capturingSignature = false;
                        }
                        else if (candidate.Contains("{"))
                        {
                            if (TryParseJavaMethodSignature(candidate, out string signature))
                            {
                                methodStack.Push(new MethodSpanBuilder
                                {
                                    Signature = signature,
                                    StartLine = signatureStartLine,
                                    StartDepth = braceDepth
                                });
                            }
                            capturingSignature = false;
                        }
                    }
                }

                braceDepth += CountChar(braceSafeLine, '{');
                braceDepth -= CountChar(braceSafeLine, '}');
                if (braceDepth < 0)
                {
                    braceDepth = 0;
                }

                while (methodStack.Count > 0 && braceDepth <= methodStack.Peek().StartDepth)
                {
                    var builder = methodStack.Pop();
                    builder.EndLine = i + 1;
                    spans.Add(builder.ToSpan());
                }
            }

            while (methodStack.Count > 0)
            {
                var builder = methodStack.Pop();
                builder.EndLine = lines.Length;
                spans.Add(builder.ToSpan());
            }

            spans.Sort((a, b) => a.StartLine.CompareTo(b.StartLine));
            return spans;
        }

        private static bool ContainsMethodToken(string trimmedLine)
        {
            if (string.IsNullOrWhiteSpace(trimmedLine))
            {
                return false;
            }

            if (trimmedLine.StartsWith("@", StringComparison.Ordinal))
            {
                return false;
            }

            int parenIndex = trimmedLine.IndexOf('(');
            if (parenIndex < 0)
            {
                return false;
            }

            string before = trimmedLine.Substring(0, parenIndex);
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

        private static int CountChar(string text, char target)
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

        private static string RemoveStringLiterals(string text)
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

        private static string RemoveComments(string line, ref bool inBlockComment)
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

        private static bool TryParseJavaMethodSignature(string candidate, out string signature)
        {
            signature = string.Empty;
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            int braceIndex = candidate.IndexOf('{');
            if (braceIndex >= 0)
            {
                candidate = candidate.Substring(0, braceIndex);
            }

            candidate = candidate.Trim();

            int parenOpen = candidate.IndexOf('(');
            int parenClose = candidate.LastIndexOf(')');
            if (parenOpen < 0 || parenClose <= parenOpen)
            {
                return false;
            }

            string before = candidate.Substring(0, parenOpen).Trim();
            if (string.IsNullOrEmpty(before) || before.Contains("="))
            {
                return false;
            }

            string parameters = candidate.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();

            var tokens = SplitTokens(before);
            tokens.RemoveAll(t => t.StartsWith("@", StringComparison.Ordinal));

            if (tokens.Count == 0)
            {
                return false;
            }

            var coreTokens = tokens.Where(t => !MethodModifiers.Contains(t)).ToList();
            if (coreTokens.Count == 0)
            {
                return false;
            }

            string methodName = coreTokens[coreTokens.Count - 1];
            if (!IsValidIdentifier(methodName) || DisallowedMethodNames.Contains(methodName))
            {
                return false;
            }

            string returnType = coreTokens.Count > 1 ? string.Join(" ", coreTokens.Take(coreTokens.Count - 1)) : string.Empty;

            string normalizedReturnType = NormalizeWhitespace(returnType);
            string normalizedParameters = NormalizeWhitespace(parameters);
            string normalizedModifiers = NormalizeWhitespace(string.Join(" ", tokens.Where(MethodModifiers.Contains)));

            var signatureParts = new List<string>();
            if (!string.IsNullOrEmpty(normalizedModifiers))
            {
                signatureParts.Add(normalizedModifiers);
            }
            if (!string.IsNullOrEmpty(normalizedReturnType))
            {
                signatureParts.Add(normalizedReturnType);
            }
            signatureParts.Add(methodName);

            signature = $"{string.Join(" ", signatureParts)}({normalizedParameters})".Trim();
            return true;
        }

        private static List<string> SplitTokens(string input)
        {
            return input.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        private static bool IsValidIdentifier(string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            if (!(char.IsLetter(token[0]) || token[0] == '_' || token[0] == '$'))
            {
                return false;
            }

            for (int i = 1; i < token.Length; i++)
            {
                char c = token[i];
                if (!(char.IsLetterOrDigit(c) || c == '_' || c == '$'))
                {
                    return false;
                }
            }

            return true;
        }

        private static string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text?.Trim() ?? string.Empty;
            }

            var sb = new StringBuilder();
            bool previousWhitespace = false;
            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!previousWhitespace)
                    {
                        sb.Append(' ');
                        previousWhitespace = true;
                    }
                }
                else
                {
                    sb.Append(c);
                    previousWhitespace = false;
                }
            }

            return sb.ToString().Trim();
        }

        private sealed class MethodSpan
        {
            public int StartLine { get; set; }
            public int EndLine { get; set; }
            public string Signature { get; set; }
        }

        private sealed class MethodSpanBuilder
        {
            public string Signature { get; set; }
            public int StartLine { get; set; }
            public int StartDepth { get; set; }
            public int EndLine { get; set; }

            public MethodSpan ToSpan()
            {
                return new MethodSpan
                {
                    StartLine = StartLine,
                    EndLine = EndLine,
                    Signature = Signature
                };
            }
        }

        private static bool StartsWithAnnotation(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return text[0] == '@';
        }
    }

    internal sealed class CSharpMethodSignatureResolver : IMethodSignatureResolver
    {
        private static readonly HashSet<string> MethodModifiers = new HashSet<string>(StringComparer.Ordinal)
        {
            "public","protected","private","internal","static","virtual","override","abstract","sealed","async","unsafe","extern","new","partial"
        };

        private static readonly HashSet<string> ControlKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "if","for","foreach","while","switch","catch","using","lock","return","else","do","checked","unchecked","fixed","try"
        };

        private static readonly HashSet<string> DisallowedMethodNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "if","for","foreach","while","switch","catch","using","lock","return","else","do","checked","unchecked","fixed","try","case","default","nameof"
        };

        private readonly List<MethodSpan> _methodSpans;

        public CSharpMethodSignatureResolver(string[] lines)
        {
            _methodSpans = BuildSpans(lines);
        }

        public string GetMethodSignature(int lineNumber)
        {
            if (_methodSpans.Count == 0)
            {
                return string.Empty;
            }

            int low = 0;
            int high = _methodSpans.Count - 1;
            while (low <= high)
            {
                int mid = (low + high) / 2;
                var span = _methodSpans[mid];
                if (lineNumber < span.StartLine)
                {
                    high = mid - 1;
                }
                else if (lineNumber > span.EndLine)
                {
                    low = mid + 1;
                }
                else
                {
                    return span.Signature;
                }
            }

            return string.Empty;
        }

        private static List<MethodSpan> BuildSpans(string[] lines)
        {
            var spans = new List<MethodSpan>();
            var methodStack = new Stack<MethodSpanBuilder>();
            bool inBlockComment = false;
            bool capturingSignature = false;
            int signatureStartLine = -1;
            var signatureBuilder = new StringBuilder();
            int braceDepth = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                string rawLine = lines[i] ?? string.Empty;
                string withoutComments = RemoveComments(rawLine, ref inBlockComment);
                string braceSafeLine = RemoveStringLiterals(withoutComments);
                string trimmed = withoutComments.Trim();

                if (!capturingSignature && ContainsMethodToken(trimmed))
                {
                    capturingSignature = true;
                    signatureBuilder.Clear();
                    signatureBuilder.Append(trimmed);
                    signatureStartLine = i + 1;
                }
                else if (capturingSignature && !string.IsNullOrWhiteSpace(trimmed))
                {
                    signatureBuilder.Append(" ").Append(trimmed);
                }

                if (capturingSignature)
                {
                    string candidate = signatureBuilder.ToString();
                    bool hasBrace = candidate.IndexOf('{') >= 0;
                    bool hasArrow = ContainsLambdaArrow(candidate);

                    if (candidate.IndexOf(';') >= 0 && !hasBrace && !hasArrow)
                    {
                        capturingSignature = false;
                    }
                    else if (hasBrace || hasArrow)
                    {
                        if (TryParseCSharpMethodSignature(candidate, out string signature, out bool hasBlockBody, out bool expressionBody))
                        {
                            if (hasBlockBody)
                            {
                                methodStack.Push(new MethodSpanBuilder
                                {
                                    Signature = signature,
                                    StartLine = signatureStartLine,
                                    StartDepth = braceDepth
                                });
                            }
                            else if (expressionBody)
                            {
                                spans.Add(new MethodSpan
                                {
                                    Signature = signature,
                                    StartLine = signatureStartLine,
                                    EndLine = i + 1
                                });
                            }
                        }
                        capturingSignature = false;
                    }
                }

                braceDepth += CountChar(braceSafeLine, '{');
                braceDepth -= CountChar(braceSafeLine, '}');
                if (braceDepth < 0)
                {
                    braceDepth = 0;
                }

                while (methodStack.Count > 0 && braceDepth <= methodStack.Peek().StartDepth)
                {
                    var builder = methodStack.Pop();
                    builder.EndLine = i + 1;
                    spans.Add(builder.ToSpan());
                }
            }

            while (methodStack.Count > 0)
            {
                var builder = methodStack.Pop();
                builder.EndLine = lines.Length;
                spans.Add(builder.ToSpan());
            }

            spans.Sort((a, b) => a.StartLine.CompareTo(b.StartLine));
            return spans;
        }

        private static bool ContainsMethodToken(string trimmedLine)
        {
            if (string.IsNullOrWhiteSpace(trimmedLine))
            {
                return false;
            }

            string normalized = StripLeadingAttributes(trimmedLine);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (normalized.StartsWith("#", StringComparison.Ordinal))
            {
                return false;
            }

            int parenIndex = normalized.IndexOf('(');
            if (parenIndex < 0)
            {
                return false;
            }

            string before = normalized.Substring(0, parenIndex);
            if (string.IsNullOrWhiteSpace(before))
            {
                return false;
            }

            string lowerBefore = before.ToLowerInvariant();
            if (lowerBefore.Contains(" class ") ||
                lowerBefore.Contains(" struct ") ||
                lowerBefore.Contains(" interface ") ||
                lowerBefore.Contains(" record ") ||
                lowerBefore.Contains(" enum ") ||
                lowerBefore.Contains(" delegate "))
            {
                return false;
            }

            string trimmedBefore = before.TrimStart();
            foreach (var keyword in ControlKeywords)
            {
                if (trimmedBefore.StartsWith(keyword + " ", StringComparison.Ordinal) ||
                    trimmedBefore.StartsWith(keyword + "\t", StringComparison.Ordinal) ||
                    trimmedBefore.StartsWith(keyword + "(", StringComparison.Ordinal) ||
                    trimmedBefore.Equals(keyword, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseCSharpMethodSignature(string candidate, out string signature, out bool hasBlockBody, out bool expressionBody)
        {
            signature = string.Empty;
            hasBlockBody = false;
            expressionBody = false;

            if (string.IsNullOrWhiteSpace(candidate))
            {
                return false;
            }

            int braceIndex = candidate.IndexOf('{');
            int arrowIndex = IndexOfLambdaArrow(candidate);
            if (arrowIndex >= 0 && (braceIndex < 0 || arrowIndex < braceIndex))
            {
                expressionBody = true;
                hasBlockBody = false;
            }
            else if (braceIndex >= 0)
            {
                hasBlockBody = true;
            }

            string working = candidate;
            if (expressionBody && arrowIndex >= 0)
            {
                working = working.Substring(0, arrowIndex);
            }
            else if (hasBlockBody && braceIndex >= 0)
            {
                working = working.Substring(0, braceIndex);
            }

            working = working.Trim();

            int parenOpen = working.IndexOf('(');
            if (parenOpen < 0)
            {
                return false;
            }

            int parenClose = FindMatchingParenthesis(working, parenOpen);
            if (parenClose < 0)
            {
                return false;
            }

            string before = working.Substring(0, parenOpen).Trim();
            string parameters = working.Substring(parenOpen + 1, parenClose - parenOpen - 1).Trim();
            string after = working.Substring(parenClose + 1).Trim();

            string whereClause = string.Empty;
            if (!string.IsNullOrEmpty(after))
            {
                int whereIndex = IndexOfWhereClause(after);
                if (whereIndex >= 0)
                {
                    whereClause = after.Substring(whereIndex).Trim();
                }
            }

            string strippedBefore = StripLeadingAttributes(before).Trim();
            if (string.IsNullOrEmpty(strippedBefore))
            {
                return false;
            }

            var tokens = SplitTokens(strippedBefore);
            if (tokens.Count == 0)
            {
                return false;
            }

            var coreTokens = tokens.Where(t => !MethodModifiers.Contains(t)).ToList();
            if (coreTokens.Count == 0)
            {
                return false;
            }

            int operatorIndex = coreTokens.IndexOf("operator");
            bool hasConversionKeyword = operatorIndex > 0 && (coreTokens[operatorIndex - 1] == "implicit" || coreTokens[operatorIndex - 1] == "explicit");
            int nameStartIndex = operatorIndex >= 0
                ? (hasConversionKeyword ? operatorIndex - 1 : operatorIndex)
                : coreTokens.Count - 1;

            if (nameStartIndex < 0 || nameStartIndex >= coreTokens.Count)
            {
                return false;
            }

            var methodNameParts = coreTokens.Skip(nameStartIndex).ToList();
            if (methodNameParts.Count == 0)
            {
                return false;
            }

            string methodName = string.Join(" ", methodNameParts);
            if (operatorIndex < 0)
            {
                if (DisallowedMethodNames.Contains(methodName) || !IsValidMethodName(methodName))
                {
                    return false;
                }
            }

            var returnTokens = coreTokens.Take(nameStartIndex).ToList();
            string returnType = string.Join(" ", returnTokens);

            string normalizedReturnType = NormalizeWhitespace(returnType);
            string normalizedParameters = NormalizeWhitespace(parameters);

            var signatureParts = new List<string>();
            if (!string.IsNullOrEmpty(normalizedReturnType))
            {
                signatureParts.Add(normalizedReturnType);
            }
            signatureParts.Add(methodName);

            string signatureText = $"{string.Join(" ", signatureParts)}({normalizedParameters})".Trim();
            if (!string.IsNullOrEmpty(whereClause))
            {
                signatureText = $"{signatureText} {NormalizeWhitespace(whereClause)}";
            }

            signature = signatureText;
            return true;
        }

        private static string StripLeadingAttributes(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string working = text.TrimStart();
            while (working.StartsWith("[", StringComparison.Ordinal))
            {
                int end = FindMatchingBracket(working, 0);
                if (end < 0)
                {
                    break;
                }

                working = working.Substring(end + 1).TrimStart();
            }
            return working;
        }

        private static int FindMatchingBracket(string text, int startIndex)
        {
            int depth = 0;
            for (int i = startIndex; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '[')
                {
                    depth++;
                }
                else if (c == ']')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private static int FindMatchingParenthesis(string text, int openIndex)
        {
            int depth = 0;
            for (int i = openIndex; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private static int IndexOfWhereClause(string text)
        {
            int index = text.IndexOf("where", StringComparison.Ordinal);
            while (index >= 0)
            {
                bool beforeOk = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                int afterIndex = index + "where".Length;
                bool afterOk = afterIndex >= text.Length || char.IsWhiteSpace(text[afterIndex]);
                if (beforeOk && afterOk)
                {
                    return index;
                }
                index = text.IndexOf("where", index + 1, StringComparison.Ordinal);
            }
            return -1;
        }

        private static bool ContainsLambdaArrow(string text)
        {
            return IndexOfLambdaArrow(text) >= 0;
        }

        private static int IndexOfLambdaArrow(string text)
        {
            return text.IndexOf("=>", StringComparison.Ordinal);
        }

        private static int CountChar(string text, char target)
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

        private static string RemoveStringLiterals(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            bool inString = false;
            char delimiter = '\0';
            bool verbatim = false;

            for (int i = 0; i < text.Length; i++)
            {
                char current = text[i];
                if (!inString)
                {
                    if (current == '"' || current == '\'')
                    {
                        inString = true;
                        delimiter = current;
                        verbatim = false;
                        if (current == '"')
                        {
                            bool prevAt = i > 0 && text[i - 1] == '@';
                            bool prevDollar = i > 0 && text[i - 1] == '$';
                            bool prevAtBefore = i > 1 && text[i - 2] == '@' && prevDollar;
                            bool prevDollarBefore = i > 1 && text[i - 2] == '$' && prevAt;
                            if (prevAt || prevAtBefore)
                            {
                                verbatim = true;
                            }
                        }
                        sb.Append(' ');
                    }
                    else
                    {
                        sb.Append(current);
                    }
                }
                else
                {
                    sb.Append(' ');
                    if (current == delimiter)
                    {
                        if (delimiter == '"' && verbatim)
                        {
                            if (i + 1 < text.Length && text[i + 1] == '"')
                            {
                                sb.Append(' ');
                                i++;
                            }
                            else
                            {
                                inString = false;
                                verbatim = false;
                                delimiter = '\0';
                            }
                        }
                        else if (i == 0 || text[i - 1] != '\\')
                        {
                            inString = false;
                            delimiter = '\0';
                        }
                    }
                }
            }

            return sb.ToString();
        }

        private static string RemoveComments(string line, ref bool inBlockComment)
        {
            if (string.IsNullOrEmpty(line))
            {
                return string.Empty;
            }

            var sb = new StringBuilder();
            bool inString = false;
            char stringDelimiter = '\0';
            bool verbatimString = false;

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
                    if (current == stringDelimiter)
                    {
                        if (stringDelimiter == '"' && verbatimString)
                        {
                            if (i + 1 < line.Length && line[i + 1] == '"')
                            {
                                sb.Append(line[i + 1]);
                                i++;
                            }
                            else
                            {
                                inString = false;
                                verbatimString = false;
                                stringDelimiter = '\0';
                            }
                        }
                        else if (i == 0 || line[i - 1] != '\\')
                        {
                            inString = false;
                            stringDelimiter = '\0';
                        }
                    }
                    continue;
                }

                if (current == '"' || current == '\'')
                {
                    inString = true;
                    stringDelimiter = current;
                    verbatimString = false;
                    if (current == '"')
                    {
                        bool prevAt = i > 0 && line[i - 1] == '@';
                        bool prevDollar = i > 0 && line[i - 1] == '$';
                        bool prevAtBefore = i > 1 && line[i - 2] == '@' && prevDollar;
                        bool prevDollarBefore = i > 1 && line[i - 2] == '$' && prevAt;
                        if (prevAt || prevAtBefore)
                        {
                            verbatimString = true;
                        }
                    }
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

        private static List<string> SplitTokens(string input)
        {
            return input.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        private static bool IsValidMethodName(string token)
        {
            if (string.IsNullOrEmpty(token))
            {
                return false;
            }

            string working = token.Trim();
            if (working.StartsWith("@", StringComparison.Ordinal))
            {
                working = working.Substring(1);
            }

            if (string.IsNullOrEmpty(working))
            {
                return false;
            }

            if (working.StartsWith("~", StringComparison.Ordinal))
            {
                return IsValidIdentifier(working.Substring(1));
            }

            if (working.Contains("."))
            {
                var parts = working.Split('.');
                return parts.All(IsValidIdentifierWithGenerics);
            }

            return IsValidIdentifierWithGenerics(working);
        }

        private static bool IsValidIdentifierWithGenerics(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            int genericIndex = text.IndexOf('<');
            if (genericIndex >= 0)
            {
                if (!text.EndsWith(">", StringComparison.Ordinal))
                {
                    return false;
                }
                string prefix = text.Substring(0, genericIndex);
                return IsValidIdentifier(prefix);
            }

            return IsValidIdentifier(text);
        }

        private static bool IsValidIdentifier(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (i == 0)
                {
                    if (!(char.IsLetter(c) || c == '_'))
                    {
                        return false;
                    }
                }
                else
                {
                    if (!(char.IsLetterOrDigit(c) || c == '_'))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text?.Trim() ?? string.Empty;
            }

            var sb = new StringBuilder();
            bool previousWhitespace = false;
            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!previousWhitespace)
                    {
                        sb.Append(' ');
                        previousWhitespace = true;
                    }
                }
                else
                {
                    sb.Append(c);
                    previousWhitespace = false;
                }
            }

            return sb.ToString().Trim();
        }

        private sealed class MethodSpan
        {
            public int StartLine { get; set; }
            public int EndLine { get; set; }
            public string Signature { get; set; }
        }

        private sealed class MethodSpanBuilder
        {
            public string Signature { get; set; }
            public int StartLine { get; set; }
            public int StartDepth { get; set; }
            public int EndLine { get; set; }

            public MethodSpan ToSpan()
            {
                return new MethodSpan
                {
                    StartLine = StartLine,
                    EndLine = EndLine,
                    Signature = Signature
                };
            }
        }
    }
}
