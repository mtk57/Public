using System.Runtime.Serialization;
using System.Collections.Generic;

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

        [DataMember]
        public List<string> FileHistory { get; set; } = new List<string>();
    }
}
