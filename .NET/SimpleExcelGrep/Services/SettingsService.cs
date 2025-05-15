using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using SimpleExcelGrep.Models;

namespace SimpleExcelGrep.Services
{
    /// <summary>
    /// アプリケーション設定の保存と読み込みを担当するサービス
    /// </summary>
    internal class SettingsService
    {
        private readonly string _settingsFilePath;
        private readonly LogService _logger;
        private const int MaxHistoryItems = 10;

        /// <summary>
        /// SettingsServiceのコンストラクタ
        /// </summary>
        /// <param name="logger">ログサービス</param>
        /// <param name="settingsFilePath">設定ファイルのパス</param>
        public SettingsService(LogService logger, string settingsFilePath = "settings.json")
        {
            _logger = logger;
            _settingsFilePath = settingsFilePath;
        }

        /// <summary>
        /// 設定ファイルからアプリケーション設定を読み込む
        /// </summary>
        /// <returns>読み込まれた設定、または新しい設定オブジェクト</returns>
        public Settings LoadSettings()
        {
            Settings settings = new Settings();
            _logger.LogMessage("設定の読み込みを開始");

            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    // DataContractJsonSerializerを使用してJSON読み込み
                    using (FileStream fs = new FileStream(_settingsFilePath, FileMode.Open))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                        settings = (Settings)serializer.ReadObject(fs);
                    }
                    _logger.LogMessage("設定の読み込みが成功しました");
                }
                else
                {
                    _logger.LogMessage("設定ファイルが見つかりません");
                    // デフォルト値を設定ファイルに保存
                    SaveSettings(settings);
                }
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"設定の読み込みに失敗: {ex.Message}");
            }

            return settings;
        }

        /// <summary>
        /// アプリケーション設定をファイルに保存
        /// </summary>
        /// <param name="settings">保存する設定オブジェクト</param>
        /// <returns>保存が成功したかどうか</returns>
        public bool SaveSettings(Settings settings)
        {
            _logger.LogMessage("設定の保存を開始");
            try
            {
                // DataContractJsonSerializerを使用してJSON保存
                using (MemoryStream ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Settings));
                    serializer.WriteObject(ms, settings);
                    File.WriteAllBytes(_settingsFilePath, ms.ToArray());
                }

                _logger.LogMessage("設定の保存が成功しました");
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogMessage($"設定の保存に失敗: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 履歴コンボボックスに項目を追加（重複なしで先頭に配置）
        /// </summary>
        public void AddToComboBoxHistory(ComboBox comboBox, string item)
        {
            if (string.IsNullOrEmpty(item))
                return;

            // 既存の項目を削除（重複を防ぐ）
            if (comboBox.Items.Contains(item))
            {
                comboBox.Items.Remove(item);
            }

            // 先頭に追加
            comboBox.Items.Insert(0, item);

            // 最大履歴数を超えた場合、古い項目を削除
            while (comboBox.Items.Count > MaxHistoryItems)
            {
                comboBox.Items.RemoveAt(comboBox.Items.Count - 1);
            }

            // 現在の選択項目を設定
            comboBox.Text = item;
        }

        /// <summary>
        /// コンボボックスの内容から履歴リストを作成
        /// </summary>
        public List<string> CreateHistoryListFromComboBox(ComboBox comboBox)
        {
            List<string> history = new List<string>();

            if (!string.IsNullOrEmpty(comboBox.Text))
            {
                history.Add(comboBox.Text);
            }

            foreach (var item in comboBox.Items)
            {
                string value = item.ToString();
                if (!history.Contains(value) && history.Count < MaxHistoryItems)
                {
                    history.Add(value);
                }
            }

            return history;
        }
    }
}