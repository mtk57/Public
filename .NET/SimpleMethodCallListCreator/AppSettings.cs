using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SimpleMethodCallListCreator
{
    [DataContract]
    public class AppSettings
    {
        [DataMember]
        public List<string> RecentFilePaths { get; set; } = new List<string>();

        [DataMember]
        public List<string> RecentIgnoreKeywords { get; set; } = new List<string>();

        [DataMember]
        public bool UseRegex { get; set; }

        [DataMember]
        public string LastIgnoreKeyword { get; set; } = string.Empty;

        [DataMember]
        public bool MatchCase { get; set; }

        [DataMember]
        public IgnoreRule IgnoreRule { get; set; } = IgnoreRule.StartsWith;

        [DataMember]
        public int SelectedIgnoreKeywordIndex { get; set; } = -1;

        [DataMember]
        public List<string> RecentCallerMethods { get; set; } = new List<string>();

        [DataMember]
        public string LastCallerMethod { get; set; } = string.Empty;

        [DataMember]
        public int SelectedCallerMethodIndex { get; set; } = -1;

        [DataMember]
        public List<IgnoreConditionSetting> IgnoreConditions { get; set; } = new List<IgnoreConditionSetting>();
    }
}
