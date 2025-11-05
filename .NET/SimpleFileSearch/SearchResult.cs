using System;

namespace SimpleFileSearch
{
    public class SearchResult
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public string Extension { get; set; }
        public long Size { get; set; }
        public DateTime LastWriteTime { get; set; }
    }
}
