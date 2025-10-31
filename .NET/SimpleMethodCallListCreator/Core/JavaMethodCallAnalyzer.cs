using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public static class JavaMethodCallAnalyzer
    {
        private const bool ExcludeConstructors = true;
        private static readonly Encoding[] CandidateEncodings = CreateCandidateEncodings();
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
            ValidateJavaFile(filePath);

            var originalText = ReadAllTextWithEncoding(filePath, out _);
            var cleanedText = RemoveComments(originalText);
            var lineIndexer = new LineIndexer(originalText);

            var classes = ExtractClasses(cleanedText, lineIndexer);
            var results = new List<MethodCallDetail>();
            if (classes.Count == 0)
            {
                return results;
            }

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

        public static List<MethodDefinitionDetail> ExtractMethodDefinitions(string filePath)
        {
            ValidateJavaFile(filePath);

            var originalText = ReadAllTextWithEncoding(filePath, out _);
            var cleanedText = RemoveComments(originalText);
            var lineIndexer = new LineIndexer(originalText);
            var packageName = ExtractPackageName(cleanedText);

            var classes = ExtractClasses(cleanedText, lineIndexer);
            var results = new List<MethodDefinitionDetail>();
            if (classes.Count == 0)
            {
                return results;
            }

            foreach (var javaClass in classes)
            {
                var methods = ExtractMethods(cleanedText, javaClass, lineIndexer);
                foreach (var method in methods)
                {
                    var signature = BuildMethodSignature(cleanedText, method);
                    var lineNumber = lineIndexer.GetLineNumber(method.SignatureIndex);
                    results.Add(new MethodDefinitionDetail(filePath, packageName, javaClass.Name, signature, lineNumber));
                }
            }

            return results;
        }

        public static JavaFileStructure ExtractMethodStructures(string filePath)
        {
            ValidateJavaFile(filePath);

            var originalText = ReadAllTextWithEncoding(filePath, out var encoding);
            var cleanedText = RemoveComments(originalText);
            var lineIndexer = new LineIndexer(originalText);

            var classes = ExtractClasses(cleanedText, lineIndexer);
            var methodStructures = new List<JavaMethodStructure>();
            if (classes.Count == 0)
            {
                return new JavaFileStructure(filePath, originalText, encoding, methodStructures);
            }

            foreach (var javaClass in classes)
            {
                var methods = ExtractMethods(cleanedText, javaClass, lineIndexer);
                foreach (var method in methods)
                {
                    var signature = BuildMethodSignature(cleanedText, method);
                    var lineNumber = lineIndexer.GetLineNumber(method.SignatureIndex);
                    var calls = new List<JavaMethodCallStructure>();
                    foreach (var call in EnumerateMethodCalls(cleanedText, method, lineIndexer))
                    {
                        calls.Add(new JavaMethodCallStructure(
                            call.CalleeIdentifier,
                            call.MethodName,
                            call.Arguments,
                            call.ArgumentCount,
                            call.MethodNameStart,
                            call.CallEndIndex,
                            call.LineNumber));
                    }

                    methodStructures.Add(new JavaMethodStructure(
                        javaClass.Name,
                        method.Name,
                        signature,
                        method.SignatureStartIndex,
                        method.BodyStartIndex,
                        method.BodyEndIndex,
                        lineNumber,
                        calls));
                }
            }

            return new JavaFileStructure(filePath, originalText, encoding, methodStructures);
        }

        private static Encoding[] CreateCandidateEncodings()
        {
            var encodings = new List<Encoding>
            {
                new UTF8Encoding(false, true)
            };

            try
            {
                encodings.Add(Encoding.GetEncoding("Shift_JIS"));
            }
            catch (ArgumentException)
            {
                // Shift_JIS が利用できない環境では追加しない
            }

            var defaultEncoding = Encoding.Default;
            if (defaultEncoding != null)
            {
                var exists = encodings.Exists(
                    e => string.Equals(e.WebName, defaultEncoding.WebName, StringComparison.OrdinalIgnoreCase));
                if (!exists)
                {
                    encodings.Add(defaultEncoding);
                }
            }

            return encodings.ToArray();
        }

        private static string ReadAllTextWithEncoding(string filePath, out Encoding encoding)
        {
            foreach (var candidate in CandidateEncodings)
            {
                try
                {
                    return ReadAllTextInternal(filePath, candidate, out encoding);
                }
                catch (DecoderFallbackException)
                {
                    // 次の候補を試す
                }
                catch (ArgumentException)
                {
                    // 次の候補を試す
                }
            }

            using (var reader = new StreamReader(filePath, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                var text = reader.ReadToEnd();
                encoding = reader.CurrentEncoding ?? Encoding.UTF8;
                return text;
            }
        }

        private static string ReadAllTextInternal(string filePath, Encoding encoding, out Encoding detectedEncoding)
        {
            using (var reader = new StreamReader(filePath, encoding, detectEncodingFromByteOrderMarks: true))
            {
                var text = reader.ReadToEnd();
                detectedEncoding = reader.CurrentEncoding ?? encoding;
                return text;
            }
        }

        private static void ValidateJavaFile(string filePath)
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

        private static string ExtractPackageName(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            const string keyword = "package";
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

                var start = SkipWhitespaceForward(text, index + keyword.Length);
                if (start >= text.Length)
                {
                    return string.Empty;
                }

                var semicolon = text.IndexOf(';', start);
                if (semicolon == -1)
                {
                    return string.Empty;
                }

                var candidate = text.Substring(start, semicolon - start).Trim();
                return candidate;
            }

            return string.Empty;
        }

        private static List<JavaClassInfo> ExtractClasses(string text, LineIndexer lineIndexer)
        {
            var classes = new List<JavaClassInfo>();
            ExtractClassesRecursive(text, 0, text.Length, lineIndexer, classes);
            return classes;
        }

        private static void ExtractClassesRecursive(string text, int startIndex, int endIndex,
            LineIndexer lineIndexer, List<JavaClassInfo> classes)
        {
            const string keywordClass = "class";
            const string keywordInterface = "interface";
            var index = startIndex;
            while (index < endIndex)
            {
                var foundClass = text.IndexOf(keywordClass, index, endIndex - index, StringComparison.Ordinal);
                var foundInterface = text.IndexOf(keywordInterface, index, endIndex - index, StringComparison.Ordinal);

                var found = foundClass;
                var isInterface = false;
                if (foundClass == -1 || (foundInterface != -1 && foundInterface < foundClass))
                {
                    found = foundInterface;
                    isInterface = true;
                }

                if (found == -1)
                {
                    break;
                }

                var keyword = isInterface ? keywordInterface : keywordClass;
                if (!IsStandaloneKeyword(text, found, keyword))
                {
                    index = found + keyword.Length;
                    continue;
                }

                var nameStart = SkipWhitespaceForward(text, found + keyword.Length);
                if (nameStart >= endIndex)
                {
                    break;
                }

                var nameEnd = nameStart;
                while (nameEnd < endIndex && IsIdentifierChar(text[nameEnd]))
                {
                    nameEnd++;
                }

                if (nameEnd == nameStart)
                {
                    index = found + keyword.Length;
                    continue;
                }

                var className = text.Substring(nameStart, nameEnd - nameStart);
                var bodyStartSearch = nameEnd;
                var bodyStart = FindNextCharWithinRange(text, bodyStartSearch, '{', endIndex);
                if (bodyStart == -1)
                {
                    var line = lineIndexer.GetLineNumber(nameStart);
                    throw new JavaParseException("クラス定義の開始位置を特定できません。", line, className);
                }

                var bodyEnd = FindMatchingBrace(text, bodyStart);
                if (bodyEnd == -1 || bodyEnd > endIndex)
                {
                    var line = lineIndexer.GetLineNumber(bodyStart);
                    throw new JavaParseException("クラス定義の終端が見つかりません。", line, className);
                }

                if (!isInterface)
                {
                    classes.Add(new JavaClassInfo
                    {
                        Name = className,
                        DeclarationIndex = found,
                        BodyStartIndex = bodyStart,
                        BodyEndIndex = bodyEnd
                    });
                }

                if (bodyStart + 1 < bodyEnd)
                {
                    ExtractClassesRecursive(text, bodyStart + 1, bodyEnd, lineIndexer, classes);
                }

                index = bodyEnd + 1;
            }
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

                    var next = SkipPossibleThrowsClause(text, closeParen + 1);
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

                    if (ExcludeConstructors && string.Equals(methodName, javaClass.Name, StringComparison.Ordinal))
                    {
                        index = bodyEnd + 1;
                        braceDepth = 0;
                        continue;
                    }

                    var signatureStart = FindMethodSignatureStart(text, methodNameStart);

                    methods.Add(new JavaMethodInfo
                    {
                        Name = methodName,
                        SignatureIndex = methodNameStart,
                        SignatureStartIndex = signatureStart,
                        ParameterListStartIndex = index,
                        ParameterListEndIndex = closeParen,
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
            foreach (var call in EnumerateMethodCalls(text, method, lineIndexer))
            {
                results.Add(new MethodCallDetail(
                    filePath,
                    javaClass.Name,
                    method.Name,
                    call.CalleeIdentifier,
                    call.MethodName,
                    call.Arguments,
                    call.LineNumber));
            }

            return results;
        }

        private static IEnumerable<MethodCallParseResult> EnumerateMethodCalls(string text,
            JavaMethodInfo method, LineIndexer lineIndexer)
        {
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

                if (text[index] != '(')
                {
                    index++;
                    continue;
                }

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

                if (ReservedKeywords.Contains(methodNameInfo.MethodName) || methodNameInfo.IsConstructorCall)
                {
                    index = closeParen + 1;
                    continue;
                }

                var callText = text.Substring(methodNameInfo.MethodNameStart,
                    closeParen - methodNameInfo.MethodNameStart + 1).Trim();
                var normalizedCallText = NormalizeCallText(callText);
                var arguments = ExtractArguments(normalizedCallText);
                var argumentCount = CountArguments(arguments);
                var calleeIdentifier = methodNameInfo.Callee;
                if (string.Equals(calleeIdentifier, "this", StringComparison.Ordinal))
                {
                    calleeIdentifier = string.Empty;
                }

                var lineNumber = lineIndexer.GetLineNumber(methodNameInfo.MethodNameStart);
                yield return new MethodCallParseResult(
                    calleeIdentifier,
                    methodNameInfo.MethodName,
                    arguments,
                    argumentCount,
                    methodNameInfo.MethodNameStart,
                    closeParen,
                    lineNumber);

                index = closeParen + 1;
            }
        }

        private static string BuildMethodSignature(string text, JavaMethodInfo method)
        {
            if (method == null)
            {
                return string.Empty;
            }

            var signatureStart = method.SignatureStartIndex;
            if (signatureStart < 0 || signatureStart >= text.Length)
            {
                signatureStart = method.SignatureIndex;
            }

            var prefixLength = method.SignatureIndex - signatureStart;
            if (prefixLength < 0)
            {
                prefixLength = 0;
            }

            var prefix = prefixLength > 0
                ? text.Substring(signatureStart, prefixLength)
                : string.Empty;

            prefix = NormalizeWhitespace(RemoveLeadingAnnotations(prefix));

            var parametersSection = string.Empty;
            if (method.ParameterListStartIndex >= 0 &&
                method.ParameterListEndIndex > method.ParameterListStartIndex &&
                method.ParameterListEndIndex < text.Length)
            {
                parametersSection = text.Substring(
                    method.ParameterListStartIndex + 1,
                    method.ParameterListEndIndex - method.ParameterListStartIndex - 1);
            }

            var parameterTypes = ExtractParameterTypes(parametersSection);
            var builder = new StringBuilder();
            if (!string.IsNullOrEmpty(prefix))
            {
                builder.Append(prefix);
                builder.Append(' ');
            }

            builder.Append(method.Name);
            builder.Append('(');
            builder.Append(string.Join(", ", parameterTypes));
            builder.Append(')');
            return builder.ToString();
        }

        private static List<string> ExtractParameterTypes(string parametersSection)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(parametersSection))
            {
                return result;
            }

            var segments = SplitParameters(parametersSection);
            foreach (var segment in segments)
            {
                var cleaned = RemoveLeadingAnnotations(segment);
                cleaned = RemoveParameterModifiers(cleaned);
                cleaned = cleaned.Trim();
                if (cleaned.Length == 0)
                {
                    continue;
                }

                var type = GetParameterType(cleaned);
                type = NormalizeWhitespace(type);
                if (type.Length > 0)
                {
                    result.Add(type);
                }
            }

            return result;
        }

        private static List<string> SplitParameters(string parametersSection)
        {
            var list = new List<string>();
            var builder = new StringBuilder();
            var angleDepth = 0;
            var parenDepth = 0;

            for (var i = 0; i < parametersSection.Length; i++)
            {
                var c = parametersSection[i];
                if (c == ',' && angleDepth == 0 && parenDepth == 0)
                {
                    list.Add(builder.ToString());
                    builder.Clear();
                    continue;
                }

                builder.Append(c);
                if (c == '<')
                {
                    angleDepth++;
                }
                else if (c == '>')
                {
                    if (angleDepth > 0)
                    {
                        angleDepth--;
                    }
                }
                else if (c == '(')
                {
                    parenDepth++;
                }
                else if (c == ')')
                {
                    if (parenDepth > 0)
                    {
                        parenDepth--;
                    }
                }
            }

            if (builder.Length > 0)
            {
                list.Add(builder.ToString());
            }

            return list;
        }

        private static string RemoveParameterModifiers(string parameter)
        {
            var result = parameter ?? string.Empty;
            result = result.TrimStart();

            var modifiers = new[] { "final" };
            var updated = true;
            while (updated && result.Length > 0)
            {
                updated = false;
                foreach (var modifier in modifiers)
                {
                    if (result.StartsWith(modifier, StringComparison.Ordinal))
                    {
                        var after = result.Substring(modifier.Length);
                        if (after.Length == 0 || char.IsWhiteSpace(after[0]))
                        {
                            result = after.TrimStart();
                            updated = true;
                            break;
                        }
                    }
                }
            }

            return result;
        }

        private static string GetParameterType(string parameter)
        {
            var trimmed = parameter.Trim();
            if (trimmed.Length == 0)
            {
                return string.Empty;
            }

            var end = trimmed.Length - 1;
            while (end >= 0 && char.IsWhiteSpace(trimmed[end]))
            {
                end--;
            }

            var arraySuffixCount = 0;
            while (end >= 1 && trimmed[end] == ']' && trimmed[end - 1] == '[')
            {
                arraySuffixCount++;
                end -= 2;
                while (end >= 0 && char.IsWhiteSpace(trimmed[end]))
                {
                    end--;
                }
            }

            var nameEnd = end;
            while (nameEnd >= 0 &&
                   (char.IsLetterOrDigit(trimmed[nameEnd]) || trimmed[nameEnd] == '_' || trimmed[nameEnd] == '$'))
            {
                nameEnd--;
            }

            var typePart = trimmed;
            if (nameEnd >= 0)
            {
                typePart = trimmed.Substring(0, nameEnd + 1).TrimEnd();
                if (typePart.Length == 0)
                {
                    typePart = trimmed;
                }
            }

            if (arraySuffixCount > 0)
            {
                var builder = new StringBuilder(typePart.Length + arraySuffixCount * 2);
                builder.Append(typePart);
                for (var i = 0; i < arraySuffixCount; i++)
                {
                    builder.Append("[]");
                }

                typePart = builder.ToString();
            }

            return typePart;
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

            var isConstructorCall = false;
            if (string.IsNullOrEmpty(callee))
            {
                var keywordEnd = SkipWhitespaceBackward(text, methodNameStart - 1);
                if (keywordEnd >= lowerBound && keywordEnd >= 0 && IsIdentifierChar(text[keywordEnd]))
                {
                    var keywordStart = FindIdentifierStart(text, keywordEnd);
                    if (keywordStart >= lowerBound)
                    {
                        var keyword = text.Substring(keywordStart, keywordEnd - keywordStart + 1);
                        if (string.Equals(keyword, "new", StringComparison.Ordinal) &&
                            IsStandaloneKeyword(text, keywordStart, "new"))
                        {
                            isConstructorCall = true;
                        }
                    }
                }
            }

            return new MethodNameInfo(methodName, methodNameStart, callee, isConstructorCall);
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

        private static int SkipPossibleThrowsClause(string text, int index)
        {
            var current = SkipWhitespaceForward(text, index);
            const string throwsKeyword = "throws";
            if (current >= text.Length)
            {
                return current;
            }

            if (current + throwsKeyword.Length > text.Length ||
                !string.Equals(text.Substring(current, throwsKeyword.Length), throwsKeyword, StringComparison.Ordinal) ||
                !IsStandaloneKeyword(text, current, throwsKeyword))
            {
                return current;
            }

            current += throwsKeyword.Length;
            while (current < text.Length)
            {
                if (text[current] == '{' || text[current] == ';')
                {
                    break;
                }

                if (text[current] == '(')
                {
                    var closeParen = FindMatchingParenthesis(text, current);
                    if (closeParen == -1)
                    {
                        return text.Length;
                    }

                    current = closeParen + 1;
                    continue;
                }

                current++;
            }

            return SkipWhitespaceForward(text, current);
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

        private static int FindNextCharWithinRange(string text, int startIndex, char target, int endIndex)
        {
            var state = new StringState();
            for (var i = startIndex; i < endIndex && i < text.Length; i++)
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

        private static string ExtractArguments(string callText)
        {
            if (string.IsNullOrEmpty(callText))
            {
                return string.Empty;
            }

            var index = callText.IndexOf('(');
            if (index < 0)
            {
                return string.Empty;
            }

            return callText.Substring(index);
        }

        private static int CountArguments(string argumentsText)
        {
            if (string.IsNullOrEmpty(argumentsText))
            {
                return 0;
            }

            var trimmed = argumentsText.Trim();
            if (trimmed.Length < 2 || trimmed[0] != '(' || trimmed[trimmed.Length - 1] != ')')
            {
                return 0;
            }

            if (trimmed.Length == 2)
            {
                return 0;
            }

            var count = 1;
            var parenthesesDepth = 0;
            var bracketDepth = 0;
            var braceDepth = 0;
            var angleDepth = 0;
            var inSingleQuote = false;
            var inDoubleQuote = false;

            for (var i = 1; i < trimmed.Length - 1; i++)
            {
                var c = trimmed[i];
                if (inSingleQuote)
                {
                    if (c == '\'' && !IsEscaped(trimmed, i))
                    {
                        inSingleQuote = false;
                    }

                    continue;
                }

                if (inDoubleQuote)
                {
                    if (c == '"' && !IsEscaped(trimmed, i))
                    {
                        inDoubleQuote = false;
                    }

                    continue;
                }

                switch (c)
                {
                    case '\'':
                        inSingleQuote = true;
                        continue;
                    case '"':
                        inDoubleQuote = true;
                        continue;
                    case '(':
                        parenthesesDepth++;
                        continue;
                    case ')':
                        if (parenthesesDepth > 0)
                        {
                            parenthesesDepth--;
                        }

                        continue;
                    case '[':
                        bracketDepth++;
                        continue;
                    case ']':
                        if (bracketDepth > 0)
                        {
                            bracketDepth--;
                        }

                        continue;
                    case '{':
                        braceDepth++;
                        continue;
                    case '}':
                        if (braceDepth > 0)
                        {
                            braceDepth--;
                        }

                        continue;
                    case '<':
                        angleDepth++;
                        continue;
                    case '>':
                        if (angleDepth > 0)
                        {
                            angleDepth--;
                        }

                        continue;
                    case ',':
                        if (parenthesesDepth == 0 &&
                            bracketDepth == 0 &&
                            braceDepth == 0 &&
                            angleDepth == 0)
                        {
                            count++;
                        }

                        continue;
                }
            }

            return count;
        }

        private static string NormalizeCallText(string callText)
        {
            if (string.IsNullOrEmpty(callText))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(callText.Length);
            var skipWhitespace = false;

            for (var i = 0; i < callText.Length; i++)
            {
                var c = callText[i];
                if (c == '\r' || c == '\n' || c == '\t')
                {
                    skipWhitespace = true;
                    continue;
                }

                if (skipWhitespace)
                {
                    if (!char.IsWhiteSpace(c))
                    {
                        builder.Append(' ');
                        skipWhitespace = false;
                    }
                    else
                    {
                        continue;
                    }
                }

                builder.Append(c);
            }

            return builder.ToString();
        }

        private static string NormalizeWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(text.Length);
            var previousIsSpace = false;
            foreach (var c in text)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!previousIsSpace)
                    {
                        builder.Append(' ');
                        previousIsSpace = true;
                    }
                }
                else
                {
                    builder.Append(c);
                    previousIsSpace = false;
                }
            }

            return builder.ToString().Trim();
        }

        private static string RemoveLeadingAnnotations(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var index = 0;
            while (index < text.Length)
            {
                index = SkipWhitespaceForward(text, index);
                if (index >= text.Length || text[index] != '@')
                {
                    break;
                }

                index++;
                while (index < text.Length &&
                       (char.IsLetterOrDigit(text[index]) || text[index] == '_' || text[index] == '.'))
                {
                    index++;
                }

                if (index < text.Length && text[index] == '(')
                {
                    var depth = 1;
                    index++;
                    while (index < text.Length && depth > 0)
                    {
                        if (text[index] == '(')
                        {
                            depth++;
                        }
                        else if (text[index] == ')')
                        {
                            depth--;
                        }

                        index++;
                    }
                }

                while (index < text.Length && char.IsWhiteSpace(text[index]))
                {
                    index++;
                }
            }

            if (index >= text.Length)
            {
                return string.Empty;
            }

            return text.Substring(index);
        }

        private static int FindMethodSignatureStart(string text, int methodNameStart)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var index = methodNameStart - 1;
            while (index >= 0)
            {
                var c = text[index];
                if (c == ';' || c == '{' || c == '}')
                {
                    index++;
                    break;
                }

                index--;
            }

            if (index < 0)
            {
                index = 0;
            }

            return SkipWhitespaceForward(text, index);
        }

        private sealed class JavaClassInfo
        {
            public string Name { get; set; }
            public int DeclarationIndex { get; set; }
            public int BodyStartIndex { get; set; }
            public int BodyEndIndex { get; set; }
        }

        private sealed class MethodCallParseResult
        {
            public MethodCallParseResult(string calleeIdentifier, string methodName, string arguments,
                int argumentCount, int methodNameStart, int callEndIndex, int lineNumber)
            {
                CalleeIdentifier = calleeIdentifier ?? string.Empty;
                MethodName = methodName;
                Arguments = arguments ?? string.Empty;
                ArgumentCount = argumentCount;
                MethodNameStart = methodNameStart;
                CallEndIndex = callEndIndex;
                LineNumber = lineNumber;
            }

            public string CalleeIdentifier { get; }
            public string MethodName { get; }
            public string Arguments { get; }
            public int ArgumentCount { get; }
            public int MethodNameStart { get; }
            public int CallEndIndex { get; }
            public int LineNumber { get; }
        }

        private sealed class JavaMethodInfo
        {
            public string Name { get; set; }
            public int SignatureIndex { get; set; }
            public int SignatureStartIndex { get; set; }
            public int ParameterListStartIndex { get; set; }
            public int ParameterListEndIndex { get; set; }
            public int BodyStartIndex { get; set; }
            public int BodyEndIndex { get; set; }
            public JavaClassInfo ClassInfo { get; set; }
        }

        private sealed class MethodNameInfo
        {
            public MethodNameInfo(string methodName, int methodNameStart, string callee, bool isConstructorCall)
            {
                MethodName = methodName;
                MethodNameStart = methodNameStart;
                Callee = callee;
                IsConstructorCall = isConstructorCall;
            }

            public string MethodName { get; }
            public int MethodNameStart { get; }
            public string Callee { get; }
            public bool IsConstructorCall { get; }
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
