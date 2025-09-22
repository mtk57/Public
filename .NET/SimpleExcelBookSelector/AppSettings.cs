using System.Runtime.Serialization;
using System.Collections.Generic;

namespace SimpleExcelBookSelector
{
    [DataContract]
    public class HistoryItem
    {
        [DataMember]
        public string FilePath { get; set; }

        [DataMember]
        public bool IsPinned { get; set; }
    }

    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public bool IsSheetSelectionEnabled { get; set; } = true;

        [DataMember]
        public bool IsAutoRefreshEnabled { get; set; } = true;

        [DataMember]
        public int RefreshInterval { get; set; } = 1;

        [DataMember]
        public List<HistoryItem> FileHistory { get; set; } = new List<HistoryItem>();

        [DataMember]
        public bool IsOpenFolderOnDoubleClickEnabled { get; set; } = true;
    }
}