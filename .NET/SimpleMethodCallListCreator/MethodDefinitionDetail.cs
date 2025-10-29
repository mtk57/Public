using System;
using System.IO;

namespace SimpleMethodCallListCreator
{
    public class MethodDefinitionDetail
    {
        public MethodDefinitionDetail(string filePath, string packageName, string className, string methodSignature)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            PackageName = packageName ?? string.Empty;
            ClassName = className ?? string.Empty;
            MethodSignature = methodSignature ?? string.Empty;
        }

        public string FilePath { get; }

        public string FileName
        {
            get { return Path.GetFileName(FilePath); }
        }

        public string PackageName { get; }

        public string ClassName { get; }

        public string MethodSignature { get; }
    }
}
