using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class JavaMethodCallAnalyzer
    {
        private static readonly HashSet<string> ReservedKeywords = new HashSet<string>(StringComparer.Ordinal)
        {
            "if",
            "for",
            "while",
            "switch",
            "catch",
            "return",
            "throw",
            "new",
            "class",
            "else",
            "case",
            "default",
            "synchronized",
            "do",
            "try"
        };

        public static List<MethodCallDetail> Analyze(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("ファイルパスが指定されていません。", nameof(filePath));
            }

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("指定されたファイルが存在しません。", filePath);
            }

            if (!string.Equals(Path.GetExtension(filePath), ".java", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("サポートされていないファイル形式です。Javaファイル（*.java）のみ指定してください。");
            }

            var originalText = File.ReadAllText(filePath, Encoding.UTF8);
            var cleanedText = RemoveComments(originalText);
            var lineIndexer = new LineIndexer(originalText);

            var classes = ExtractClasses(cleanedText, lineIndexer);
            if (classes.Count == 0)
            {
                throw new JavaParseException("クラス定義が見つかりません。", 1);
            }

            var results = new List<MethodCallDetail>();
            foreach (var javaClass in classes)
            {
                var methods = ExtractMethods(cleanedText, javaClass, lineIndexer);
                foreach (var method in methods)
                {
                    results.AddRange(ExtractMethodCalls(cleanedText, javaClass, method, lineIndexer, filePath));
                }
            }

            return results;
        }

        private static string RemoveComments(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var chars = text.ToCharArray();
            var inLineComment = false;
            var inBlockComment = false;
            var inDoubleQuote = false;
            var inSingleQuote = false;

            for (var i = 0; i < chars.Length; i++)
            {
                if (inLineComment)
                {
                    if (chars[i] == '\n')
                    {
                        inLineComment = false;
                    }
                    else if (chars[i] != '\r')
                    {
                        chars[i] = ' ';
                    }
                    continue;
                }

                if (inBlockComment)
                {
                    if (chars[i] == '*' && i + 1 < chars.Length && chars[i + 1] == '/')
                    {
                        chars[i] = ' ';
                        chars[i + 1] = ' ';
                        inBlockComment = false;
                        i++;
                    }
                    else if (chars[i] != '\r' && chars[i] != '\n')
                    {
                        chars[i] = ' ';
                    }
                    continue;
                }

                if (!inSingleQuote && chars[i] == '"' && !IsEscaped(chars, i))
                {
                    inDoubleQuote = !inDoubleQuote;
                    continue;
                }

                if (!inDoubleQuote && chars[i] == '\'' && !IsEscaped(chars, i))
                {
                    inSingleQuote = !inSingleQuote;
                    continue;
                }

                if (inDoubleQuote || inSingleQuote)
                {
                    continue;
                }

                if (chars[i] == '/' && i + 1 < chars.Length)
                {
                    if (chars[i + 1] == '/')
                    {
                        chars[i] = ' ';
                        chars[i + 1] = ' ';
                        inLineComment = true;
                        i++;
                    }
                    else if (chars[i + 1] == '*')
                    {
                        chars[i] = ' ';
                        chars[i + 1] = ' ';
                        inBlockComment = true;
                        i++;
                    }
                }
            }

            return new string(chars);
        }

        private static List<JavaClassInfo> ExtractClasses(string text, LineIndexer lineIndexer)
        {
            var classes = new List<JavaClassInfo>();
            var keyword = "class";
            var index = 0;
            while (index < text.Length)
            {
                index = text.IndexOf(keyword, index, StringComparison.Ordinal);
                if (index == -1)
                {
                    break;
                }

                if (!IsStandaloneKeyword(text, index, keyword))
                {
                    index += keyword.Length;
                    continue;
                }

                var nameStart = SkipWhitespaceForward(text, index + keyword.Length);
                if (nameStart >= text.Length)
                {
                    break;
                }

                var nameEnd = nameStart;
                while (nameEnd < text.Length && IsIdentifierChar(text[nameEnd]))
                {
                    nameEnd++;
                }

                if (nameEnd == nameStart)
                {
                    index += keyword.Length;
                    continue;
                }

                var className = text.Substring(nameStart, nameEnd - nameStart);
                var bodyStartSearch = nameEnd;
                var bodyStart = FindNextChar(text, bodyStartSearch, '{');
                if (bodyStart == -1)
                {
                    var line = lineIndexer.GetLineNumber(nameStart);
                    throw new JavaParseException("クラス定義の開始位置を特定できません。", line, className);
                }

                var bodyEnd = FindMatchingBrace(text, bodyStart);
                if (bodyEnd == -1)
                {
                    var line = lineIndexer.GetLineNumber(bodyStart);
                    throw new JavaParseException("クラス定義の終端が見つかりません。", line, className);
                }

                classes.Add(new JavaClassInfo
                {
                    Name = className,
                    DeclarationIndex = index,
                    BodyStartIndex = bodyStart,
                    BodyEndIndex = bodyEnd
                });

                index = bodyEnd + 1;
            }

            return classes;
        }

        private static List<JavaMethodInfo> ExtractMethods(string text, JavaClassInfo javaClass, LineIndexer lineIndexer)
        {
            var methods = new List<JavaMethodInfo>();
            var index = javaClass.BodyStartIndex + 1;
            var classBodyEnd = javaClass.BodyEndIndex;
            var braceDepth = 0;
            var state = new StringState();

            while (index < classBodyEnd)
            {
                UpdateStringState(text, index, state);
                if (state.InString)
                {
                    index++;
                    continue;
                }

                var current = text[index];
                if (current == '{')
                {
                    braceDepth++;
                }
                else if (current == '}')
                {
                    if (braceDepth > 0)
                    {
                        braceDepth--;
                    }
                }
                else if (current == '(' && braceDepth == 0)
                {
                    var closeParen = FindMatchingParenthesis(text, index);
                    if (closeParen == -1)
                    {
                        var line = lineIndexer.GetLineNumber(index);
                        throw new JavaParseException("メソッド定義の括弧が閉じられていません。", line);
                    }

                    var next = SkipWhitespaceForward(text, closeParen + 1);
                    if (next >= text.Length || text[next] != '{')
                    {
                        index++;
                        continue;
                    }

                    var methodNameEnd = FindIdentifierEndBackward(text, index - 1);
                    if (methodNameEnd < 0)
                    {
                        index++;
                        continue;
                    }

                    var methodNameStart = FindIdentifierStart(text, methodNameEnd);
                    if (methodNameStart < 0)
                    {
                        index++;
                        continue;
                    }

                    var methodName = text.Substring(methodNameStart, methodNameEnd - methodNameStart + 1);
                    if (string.IsNullOrWhiteSpace(methodName))
                    {
                        index++;
                        continue;
                    }

                    var bodyEnd = FindMatchingBrace(text, next);
                    if (bodyEnd == -1)
                    {
                        var line = lineIndexer.GetLineNumber(next);
                        throw new JavaParseException("メソッド定義の終端が見つかりません。", line, methodName);
                    }

                    methods.Add(new JavaMethodInfo
                    {
                        Name = methodName,
                        SignatureIndex = methodNameStart,
                        BodyStartIndex = next,
                        BodyEndIndex = bodyEnd,
                        ClassInfo = javaClass
                    });

                    index = bodyEnd + 1;
                    braceDepth = 0;
                    continue;
                }

                index++;
            }

            return methods;
        }

        private static List<MethodCallDetail> ExtractMethodCalls(string text, JavaClassInfo javaClass,
            JavaMethodInfo method, LineIndexer lineIndexer, string filePath)
        {
            var results = new List<MethodCallDetail>();
            var state = new StringState();
            var index = method.BodyStartIndex + 1;
            var end = method.BodyEndIndex;

            while (index < end)
            {
                UpdateStringState(text, index, state);
                if (state.InString)
                {
                    index++;
                    continue;
                }

                if (text[index] == '(')
                {
                    var closeParen = FindMatchingParenthesis(text, index);
                    if (closeParen == -1)
                    {
                        var line = lineIndexer.GetLineNumber(index);
                        throw new JavaParseException("メソッド呼び出しの括弧が閉じられていません。", line, method.Name);
                    }

                    var methodNameInfo = GetMethodNameInfo(text, method.BodyStartIndex + 1, index - 1);
                    if (methodNameInfo == null)
                    {
                        index = closeParen + 1;
                        continue;
                    }

                    if (ReservedKeywords.Contains(methodNameInfo.MethodName))
                    {
                        index = closeParen + 1;
                        continue;
                    }

                    var callText = text.Substring(methodNameInfo.MethodNameStart,
                        closeParen - methodNameInfo.MethodNameStart + 1).Trim();

                    var calleeClass = methodNameInfo.Callee;
                    if (string.Equals(calleeClass, "this", StringComparison.Ordinal))
                    {
                        calleeClass = string.Empty;
                    }

                    var lineNumber = lineIndexer.GetLineNumber(methodNameInfo.MethodNameStart);
                    results.Add(new MethodCallDetail(
                        filePath,
                        javaClass.Name,
                        method.Name,
                        calleeClass,
                        callText,
                        lineNumber));

                    index = closeParen + 1;
                    continue;
                }

                index++;
            }

            return results;
        }

        private static MethodNameInfo GetMethodNameInfo(string text, int lowerBound, int searchIndex)
        {
            var methodNameEnd = FindIdentifierEndBackward(text, searchIndex);
            if (methodNameEnd < lowerBound)
            {
                return null;
            }

            if (methodNameEnd < 0)
            {
                return null;
            }

            var methodNameStart = FindIdentifierStart(text, methodNameEnd);
            if (methodNameStart < lowerBound)
            {
                return null;
            }

            if (methodNameStart > methodNameEnd)
            {
                return null;
            }

            var methodName = text.Substring(methodNameStart, methodNameEnd - methodNameStart + 1);
            if (string.IsNullOrEmpty(methodName))
            {
                return null;
            }

            var callee = string.Empty;
            var preIndex = SkipWhitespaceBackward(text, methodNameStart - 1);
            if (preIndex >= lowerBound && preIndex >= 0 && text[preIndex] == '.')
            {
                var calleeEnd = SkipWhitespaceBackward(text, preIndex - 1);
                if (calleeEnd < lowerBound || calleeEnd < 0 || !IsIdentifierChar(text[calleeEnd]))
                {
                    return null;
                }

                var calleeStart = FindIdentifierStart(text, calleeEnd);
                if (calleeStart < lowerBound)
                {
                    return null;
                }

                var preceding = SkipWhitespaceBackward(text, calleeStart - 1);
                if (preceding >= lowerBound && preceding >= 0 && text[preceding] == '.')
                {
                    return null;
                }

                callee = text.Substring(calleeStart, calleeEnd - calleeStart + 1);
            }

            return new MethodNameInfo(methodName, methodNameStart, callee);
        }

        private static bool IsStandaloneKeyword(string text, int index, string keyword)
        {
            var before = index - 1;
            if (before >= 0 && IsIdentifierChar(text[before]))
            {
                return false;
            }

            var after = index + keyword.Length;
            if (after < text.Length && IsIdentifierChar(text[after]))
            {
                return false;
            }

            return true;
        }

        private static int SkipWhitespaceForward(string text, int index)
        {
            var current = index;
            while (current < text.Length && char.IsWhiteSpace(text[current]))
            {
                current++;
            }

            return current;
        }

        private static int SkipWhitespaceBackward(string text, int index)
        {
            var current = index;
            while (current >= 0 && char.IsWhiteSpace(text[current]))
            {
                current--;
            }

            return current;
        }

        private static int FindNextChar(string text, int startIndex, char target)
        {
            var state = new StringState();
            for (var i = startIndex; i < text.Length; i++)
            {
                UpdateStringState(text, i, state);
                if (state.InString)
                {
                    continue;
                }

                if (text[i] == target)
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindMatchingBrace(string text, int startIndex)
        {
            var state = new StringState();
            var depth = 0;
            for (var i = startIndex; i < text.Length; i++)
            {
                UpdateStringState(text, i, state);
                if (state.InString)
                {
                    continue;
                }

                if (text[i] == '{')
                {
                    depth++;
                }
                else if (text[i] == '}')
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

        private static int FindMatchingParenthesis(string text, int startIndex)
        {
            var state = new StringState();
            var depth = 0;
            for (var i = startIndex; i < text.Length; i++)
            {
                UpdateStringState(text, i, state);
                if (state.InString)
                {
                    continue;
                }

                if (text[i] == '(')
                {
                    depth++;
                }
                else if (text[i] == ')')
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

        private static int FindIdentifierEndBackward(string text, int index)
        {
            var current = SkipWhitespaceBackward(text, index);
            if (current < 0)
            {
                return -1;
            }

            if (!IsIdentifierChar(text[current]))
            {
                return -1;
            }

            return current;
        }

        private static int FindIdentifierStart(string text, int index)
        {
            var current = index;
            while (current >= 0 && IsIdentifierChar(text[current]))
            {
                current--;
            }

            return current + 1;
        }

        private static bool IsIdentifierChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '$';
        }

        private static void UpdateStringState(string text, int index, StringState state)
        {
            if (text[index] == '"' && !state.InSingleQuote && !IsEscaped(text, index))
            {
                state.InDoubleQuote = !state.InDoubleQuote;
                return;
            }

            if (text[index] == '\'' && !state.InDoubleQuote && !IsEscaped(text, index))
            {
                state.InSingleQuote = !state.InSingleQuote;
                return;
            }
        }

        private static bool IsEscaped(IList<char> chars, int index)
        {
            var backslashCount = 0;
            var current = index - 1;
            while (current >= 0 && chars[current] == '\\')
            {
                backslashCount++;
                current--;
            }

            return backslashCount % 2 == 1;
        }

        private static bool IsEscaped(string text, int index)
        {
            var backslashCount = 0;
            var current = index - 1;
            while (current >= 0 && text[current] == '\\')
            {
                backslashCount++;
                current--;
            }

            return backslashCount % 2 == 1;
        }

        private sealed class JavaClassInfo
        {
            public string Name { get; set; }
            public int DeclarationIndex { get; set; }
            public int BodyStartIndex { get; set; }
            public int BodyEndIndex { get; set; }
        }

        private sealed class JavaMethodInfo
        {
            public string Name { get; set; }
            public int SignatureIndex { get; set; }
            public int BodyStartIndex { get; set; }
            public int BodyEndIndex { get; set; }
            public JavaClassInfo ClassInfo { get; set; }
        }

        private sealed class MethodNameInfo
        {
            public MethodNameInfo(string methodName, int methodNameStart, string callee)
            {
                MethodName = methodName;
                MethodNameStart = methodNameStart;
                Callee = callee;
            }

            public string MethodName { get; }
            public int MethodNameStart { get; }
            public string Callee { get; }
        }

        private sealed class StringState
        {
            public bool InSingleQuote { get; set; }
            public bool InDoubleQuote { get; set; }
            public bool InString
            {
                get { return InSingleQuote || InDoubleQuote; }
            }
        }

        private sealed class LineIndexer
        {
            private readonly int[] _lineStarts;

            public LineIndexer(string text)
            {
                var starts = new List<int> { 0 };
                for (var i = 0; i < text.Length; i++)
                {
                    if (text[i] == '\n')
                    {
                        starts.Add(i + 1);
                    }
                }

                _lineStarts = starts.ToArray();
            }

            public int GetLineNumber(int index)
            {
                if (_lineStarts.Length == 0)
                {
                    return 1;
                }

                if (index <= 0)
                {
                    return 1;
                }

                var position = Array.BinarySearch(_lineStarts, index);
                if (position >= 0)
                {
                    return position + 1;
                }

                position = ~position;
                if (position <= 0)
                {
                    return 1;
                }

                if (position > _lineStarts.Length)
                {
                    return _lineStarts.Length;
                }

                return position;
            }
        }
    }
}
