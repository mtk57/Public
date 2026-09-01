using System.Collections.Generic;
using System.Runtime.Serialization;

namespace Dir2Txt
{
    [DataContract]
    internal class AppSettings
    {
        [DataMember]
        public string DirPath { get; set; }

        [DataMember]
        public List<string> DirPathHistories { get; set; }

        [DataMember]
        public string IgnoreDirs { get; set; }

        [DataMember]
        public List<string> IgnoreDirsHistories { get; set; }

        [DataMember]
        public string IgnoreFiles { get; set; }

        [DataMember]
        public List<string> IgnoreFilesHistories { get; set; }

        [DataMember]
        public string IgnoreExt { get; set; }

        [DataMember]
        public List<string> IgnoreExtHistories { get; set; }

        [DataMember]
        public bool IgnoreExtNegated { get; set; }

        [DataMember]
        public bool OutputToFile { get; set; }

        [DataMember]
        public string DivideLength { get; set; }

        [DataMember]
        public List<string> ExtractDirPathHistories { get; set; }
    }
}
