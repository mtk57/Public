using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace SimpleSqlAdjuster
{
    internal sealed class UserSettings
    {
        public int WindowWidth { get; set; }

        public int WindowHeight { get; set; }

        public int WindowX { get; set; }

        public int WindowY { get; set; }

        public string LastBeforeSql { get; set; }
    }

    internal sealed class UserSettingsService
    {
        private const string SettingsFileName = "SimpleSqlAdjuster.settings.json";

        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        public UserSettings Load()
        {
            try
            {
                var path = GetSettingsPath();
                if (!File.Exists(path))
                {
                    return new UserSettings();
                }

                var json = File.ReadAllText(path, Encoding.UTF8);
                var settings = _serializer.Deserialize<UserSettings>(json);
                return settings ?? new UserSettings();
            }
            catch (Exception ex)
            {
                LogService.Log(ex, "設定ファイルの読み込みに失敗しました。");
                return new UserSettings();
            }
        }

        public void Save(UserSettings settings, Size windowSize, Point windowLocation)
        {
            if (settings == null)
            {
                return;
            }

            settings.WindowWidth = windowSize.Width;
            settings.WindowHeight = windowSize.Height;
            settings.WindowX = windowLocation.X;
            settings.WindowY = windowLocation.Y;

            try
            {
                var json = _serializer.Serialize(settings);
                var path = GetSettingsPath();
                File.WriteAllText(path, json, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogService.Log(ex, "設定ファイルの保存に失敗しました。");
            }
        }

        private static string GetSettingsPath()
        {
            var basePath = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(basePath, SettingsFileName);
        }
    }
}
