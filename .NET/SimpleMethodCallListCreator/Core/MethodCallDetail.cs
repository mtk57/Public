using System;
using System.IO;

namespace SimpleMethodCallListCreator
{
    public class MethodCallDetail
    {
        public MethodCallDetail(string filePath, string className, string callerMethod,
            string calleeClass, string calleeMethodName, string calleeMethodArguments, int lineNumber)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            FileName = Path.GetFileName(filePath);
            ClassName = className;
            CallerMethod = callerMethod;
            CalleeClass = calleeClass;
            CalleeMethodName = calleeMethodName;
            CalleeMethodArguments = calleeMethodArguments;
            LineNumber = lineNumber;
        }

        public string FilePath { get; }

        public string FileName { get; }

        public string ClassName { get; }

        public string CallerMethod { get; }

        public string CalleeClass { get; }

        public string CalleeMethodName { get; }

        public string CalleeMethodArguments { get; }

        public string CalleeMethodFull
        {
            get { return string.Concat(CalleeMethodName, CalleeMethodArguments); }
        }

        public int LineNumber { get; }
    }
}
