using System;
using System.IO;

namespace PPTKalem
{
    /// <summary>
    /// Uygulama ayarlarını %APPDATA%\PPTKalem\settings.ini dosyasında saklar.
    /// </summary>
    public class AppSettings
    {
        private static AppSettings _instance;
        public static AppSettings Instance => _instance ?? (_instance = Load());

        public bool AutoShowOnSlideShow { get; set; } = true;
        public bool DarkTheme           { get; set; } = true;

        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PPTKalem");
        private static readonly string SettingsFile =
            Path.Combine(SettingsDir, "settings.ini");

        public static AppSettings Load()
        {
            var s = new AppSettings();
            try
            {
                if (!File.Exists(SettingsFile)) return s;
                foreach (var line in File.ReadAllLines(SettingsFile))
                {
                    var parts = line.Split('=');
                    if (parts.Length != 2) continue;
                    string key = parts[0].Trim(), val = parts[1].Trim();
                    if      (key == "AutoShowOnSlideShow") s.AutoShowOnSlideShow = val == "True";
                    else if (key == "DarkTheme")           s.DarkTheme           = val == "True";
                }
            }
            catch { }
            return s;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);
                File.WriteAllLines(SettingsFile, new[]
                {
                    $"AutoShowOnSlideShow={AutoShowOnSlideShow}",
                    $"DarkTheme={DarkTheme}"
                });
            }
            catch { }
        }
    }
}
