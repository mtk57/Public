using System;
using System.Collections.Generic;

namespace SimpleFileSearch
{
    // 設定保存用クラス
    [Serializable]
    public class AppSettings
    {
        public List<string> KeywordHistory { get; set; } = new List<string>();
        public List<string> FolderPathHistory { get; set; } = new List<string>();
        public bool UseRegex { get; set; } = false;
        public bool IncludeFolderNames { get; set; } = false;
    }
}
