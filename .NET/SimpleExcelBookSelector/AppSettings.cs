using System.Runtime.Serialization;

namespace SimpleExcelBookSelector
{
    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public bool IsSheetSelectionEnabled { get; set; } = true;

        [DataMember]
        public bool IsAutoRefreshEnabled { get; set; } = true;

        [DataMember]
        public int RefreshInterval { get; set; } = 1;
    }
}
