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
        public bool MatchCase { get; set; }

        [DataMember]
        public IgnoreRule IgnoreRule { get; set; } = IgnoreRule.StartsWith;
    }
}
