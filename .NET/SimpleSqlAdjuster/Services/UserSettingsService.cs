using System.Drawing;

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
        public UserSettings Load()
        {
            return new UserSettings();
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
        }
    }
}
