using System;

namespace SimpleMethodCallListCreator
{
    public class JavaParseException : Exception
    {
        public JavaParseException(string message, int lineNumber, string invalidContent = null)
            : base(message)
        {
            LineNumber = lineNumber;
            InvalidContent = invalidContent;
        }

        public int LineNumber { get; }

        public string InvalidContent { get; }
    }
}
