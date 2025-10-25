using System;
using System.IO;

namespace SimpleMethodCallListCreator
{
    public class MethodCallDetail
    {
        public MethodCallDetail(string filePath, string className, string callerMethod,
            string calleeClass, string calleeMethod, int lineNumber)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            FileName = Path.GetFileName(filePath);
            ClassName = className;
            CallerMethod = callerMethod;
            CalleeClass = calleeClass;
            CalleeMethod = calleeMethod;
            LineNumber = lineNumber;
        }

        public string FilePath { get; }

        public string FileName { get; }

        public string ClassName { get; }

        public string CallerMethod { get; }

        public string CalleeClass { get; }

        public string CalleeMethod { get; }

        public int LineNumber { get; }
    }
}
