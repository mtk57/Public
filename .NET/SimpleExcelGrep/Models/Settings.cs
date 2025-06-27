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

        [DataMember]
        public double IgnoreFileSizeMB { get; set; } = 0; // 無視するファイルサイズ (MB単位)

        // ===== 機能追加 =====
        [DataMember]
        public bool CellModeEnabled { get; set; } = false; // セル検索モードが有効か

        [DataMember]
        public string CellAddress { get; set; } = ""; // 検索対象のセルアドレス

        [DataMember]
        public bool SearchSubDirectories { get; set; } = true; // サブフォルダも検索対象にするか

        [DataMember]
        public bool EnableLog { get; set; } = true; // ログファイルを出力するか

        [DataMember]
        public bool SearchInvisibleSheets { get; set; } = true; // 非表示シートも検索対象にするか
    }
}