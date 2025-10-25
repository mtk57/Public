using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace SimpleMethodCallListCreator
{
    public static class SettingsManager
    {
        private const string SettingsFileName = "settings.json";

        public static AppSettings Load()
        {
            try
            {
                var path = GetSettingsFilePath();
                if (!File.Exists(path))
                {
                    return new AppSettings();
                }

                using (var stream = File.OpenRead(path))
                {
                    var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                    var settings = serializer.ReadObject(stream) as AppSettings;
                    return settings ?? new AppSettings();
                }
            }
            catch
            {
                return new AppSettings();
            }
        }

        public static void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            var path = GetSettingsFilePath();
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (var stream = File.Create(path))
            {
                var serializer = new DataContractJsonSerializer(typeof(AppSettings));
                serializer.WriteObject(stream, settings);
            }
        }

        private static string GetSettingsFilePath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var directory = Path.Combine(appData, "SimpleMethodCallListCreator");
            return Path.Combine(directory, SettingsFileName);
        }
    }
}
