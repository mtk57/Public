using System.Runtime.Serialization;

namespace SimpleMethodCallListCreator
{
    [DataContract]
    public class IgnoreConditionSetting
    {
        [DataMember]
        public string Keyword { get; set; } = string.Empty;

        [DataMember]
        public IgnoreRule Rule { get; set; } = IgnoreRule.StartsWith;

        [DataMember]
        public bool UseRegex { get; set; }

        [DataMember]
        public bool MatchCase { get; set; }
    }
}
