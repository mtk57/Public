using System;

namespace SimpleGrep.Core
{
    internal sealed class SearchResult
    {
        public string FilePath { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public int LineNumber { get; set; }
        public string LineText { get; set; } = string.Empty;
        public string MethodSignature { get; set; } = string.Empty;
        public string EncodingName { get; set; } = string.Empty;
    }
}
