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
        public bool UsePartialMatch { get; set; } = false;
        public bool SearchSubDir { get; set; } = true;
        public bool DblClickToOpen { get; set; } = false;
        public string LastKeyword { get; set; } = string.Empty;
        public string LastFolderPath { get; set; } = string.Empty;
    }
}
