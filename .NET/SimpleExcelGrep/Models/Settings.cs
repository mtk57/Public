using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SimpleExcelGrep.Models
{
    /// <summary>
    /// アプリケーション設定を格納するモデルクラス
    /// </summary>
    [DataContract]
    internal class Settings
    {
        [DataMember]
        public string FolderPath { get; set; } = "";

        [DataMember]
        public string SearchKeyword { get; set; } = "";

        [DataMember]
        public bool UseRegex { get; set; } = false;

        [DataMember]
        public string IgnoreKeywords { get; set; } = "";

        [DataMember]
        public List<string> FolderPathHistory { get; set; } = new List<string>();

        [DataMember]
        public List<string> SearchKeywordHistory { get; set; } = new List<string>();

        [DataMember]
        public List<string> IgnoreKeywordsHistory { get; set; } = new List<string>();

        [DataMember]
        public bool RealTimeDisplay { get; set; } = true; // リアルタイム表示設定

        [DataMember]
        public int MaxParallelism { get; set; } = Environment.ProcessorCount; // 並列処理数

        [DataMember]
        public bool FirstHitOnly { get; set; } = false; // 最初のヒットのみ検索設定

        [DataMember]
        public bool SearchShapes { get; set; } = false; // 図形内の文字列を検索するかどうか
    }
}