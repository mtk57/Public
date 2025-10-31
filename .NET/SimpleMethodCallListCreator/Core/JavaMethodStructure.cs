using System;
using System.Collections.Generic;
using System.Text;

namespace SimpleMethodCallListCreator
{
    public sealed class JavaFileStructure
    {
        public JavaFileStructure(string filePath, string originalText, Encoding encoding,
            IReadOnlyList<JavaMethodStructure> methods)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            OriginalText = originalText ?? string.Empty;
            Encoding = encoding;
            Methods = methods ?? throw new ArgumentNullException(nameof(methods));
        }

        public string FilePath { get; }

        public string OriginalText { get; }

        public Encoding Encoding { get; }

        public IReadOnlyList<JavaMethodStructure> Methods { get; }

        public JavaMethodStructure FindMethodBySignature(string signature)
        {
            if (signature == null)
            {
                return null;
            }

            foreach (var method in Methods)
            {
                if (string.Equals(method.MethodSignature, signature, StringComparison.Ordinal))
                {
                    return method;
                }
            }

            return null;
        }
    }

    public sealed class JavaMethodStructure
    {
        public JavaMethodStructure(string className, string methodName, string methodSignature,
            int signatureStartIndex, int bodyStartIndex, int bodyEndIndex, int lineNumber,
            IReadOnlyList<JavaMethodCallStructure> calls)
        {
            ClassName = className ?? string.Empty;
            MethodName = methodName ?? string.Empty;
            MethodSignature = methodSignature ?? string.Empty;
            SignatureStartIndex = signatureStartIndex;
            BodyStartIndex = bodyStartIndex;
            BodyEndIndex = bodyEndIndex;
            LineNumber = lineNumber;
            Calls = calls ?? Array.Empty<JavaMethodCallStructure>();
        }

        public string ClassName { get; }

        public string MethodName { get; }

        public string MethodSignature { get; }

        public int SignatureStartIndex { get; }

        public int BodyStartIndex { get; }

        public int BodyEndIndex { get; }

        public int LineNumber { get; }

        public IReadOnlyList<JavaMethodCallStructure> Calls { get; }
    }

    public sealed class JavaMethodCallStructure
    {
        public JavaMethodCallStructure(string calleeIdentifier, string methodName, string argumentsText,
            int argumentCount, int methodNameStartIndex, int callEndIndex, int lineNumber)
        {
            CalleeIdentifier = calleeIdentifier ?? string.Empty;
            MethodName = methodName ?? string.Empty;
            ArgumentsText = argumentsText ?? string.Empty;
            ArgumentCount = argumentCount < 0 ? 0 : argumentCount;
            MethodNameStartIndex = methodNameStartIndex;
            CallEndIndex = callEndIndex;
            LineNumber = lineNumber;
        }

        public string CalleeIdentifier { get; }

        public string MethodName { get; }

        public string ArgumentsText { get; }

        public int ArgumentCount { get; }

        public int MethodNameStartIndex { get; }

        public int CallEndIndex { get; }

        public int LineNumber { get; }

        public bool HasExplicitCallee
        {
            get { return !string.IsNullOrEmpty(CalleeIdentifier); }
        }
    }
}
