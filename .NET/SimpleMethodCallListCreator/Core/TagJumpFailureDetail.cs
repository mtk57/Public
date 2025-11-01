using System;

namespace SimpleMethodCallListCreator
{
    public sealed class TagJumpFailureDetail
    {
        public TagJumpFailureDetail(string filePath, int lineNumber, string callerMethodSignature,
            string callExpression, string reason)
        {
            FilePath = filePath ?? string.Empty;
            LineNumber = lineNumber < 0 ? 0 : lineNumber;
            CallerMethodSignature = callerMethodSignature ?? string.Empty;
            CallExpression = callExpression ?? string.Empty;
            Reason = reason ?? string.Empty;
        }

        public string FilePath { get; }

        public int LineNumber { get; }

        public string CallerMethodSignature { get; }

        public string CallExpression { get; }

        public string Reason { get; }
    }
}
