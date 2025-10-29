using System;
using System.IO;

namespace SimpleMethodCallListCreator
{
    public class MethodDefinitionDetail
    {
        public MethodDefinitionDetail(string filePath, string packageName, string className, string methodSignature, int lineNumber = -1)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            PackageName = packageName ?? string.Empty;
            ClassName = className ?? string.Empty;
            MethodSignature = methodSignature ?? string.Empty;
            LineNumber = lineNumber > 0 ? lineNumber : -1;
        }

        public string FilePath { get; }

        public string FileName
        {
            get { return Path.GetFileName(FilePath); }
        }

        public string PackageName { get; }

        public string ClassName { get; }

        public string MethodSignature { get; }

        public int LineNumber { get; }
    }
}
