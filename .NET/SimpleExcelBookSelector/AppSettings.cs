using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SimpleExcelBookSelector
{
    [DataContract]
    public class HistoryItem
    {
        [DataMember]
        public string FilePath { get; set; }

        [DataMember]
        public bool IsPinned { get; set; }

        [DataMember]
        public DateTime? LastUpdated { get; set; }
    }

    [DataContract]
    public class DataGridColumnLayout
    {
        [DataMember]
        public int Width { get; set; }

        [DataMember]
        public double FillWeight { get; set; }

        [DataMember]
        public string AutoSizeMode { get; set; }
    }

    [DataContract]
    public class FormLayoutSettings
    {
        [DataMember]
        public int Width { get; set; }

        [DataMember]
        public int Height { get; set; }

        [DataMember]
        public string WindowState { get; set; }

        [DataMember]
        public Dictionary<string, DataGridColumnLayout> ColumnLayouts { get; set; } = new Dictionary<string, DataGridColumnLayout>(StringComparer.Ordinal);
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

        [DataMember]
        public FormLayoutSettings MainFormLayout { get; set; } = new FormLayoutSettings();

        [DataMember]
        public FormLayoutSettings HistoryFormLayout { get; set; } = new FormLayoutSettings();
    }
}
