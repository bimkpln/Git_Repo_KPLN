using KPLN_CommandsWheel.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_CommandsWheel.Services
{
    internal static class UserSettingsService
    {
        private const int MaxWheelCommands = 8;
        private const int MaxRecentCommands = 20;

        internal static string SettingsDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "KPLN",
                    "CommandsWheel"
                );
            }
        }

        internal static string SettingsPath
        {
            get { return Path.Combine(SettingsDirectory, "settings.json"); }
        }

        internal static UserSettings Load()
        {
            UserSettings settings = null;

            if (File.Exists(SettingsPath))
            {
                try
                {
                    settings = JsonSerialization.Deserialize<UserSettings>(File.ReadAllText(SettingsPath, Encoding.UTF8));
                }
                catch
                {
                    settings = null;
                }
            }

            if (settings == null)
            {
                settings = new UserSettings();
            }

            Normalize(settings);
            return settings;
        }

        internal static void Save(UserSettings settings)
        {
            if (settings == null)
            {
                return;
            }

            Normalize(settings);
            Directory.CreateDirectory(SettingsDirectory);
            string json = JsonSerialization.Serialize(settings);
            File.WriteAllText(SettingsPath, json, Encoding.UTF8);
        }

        internal static void AddRecent(UserSettings settings, string commandId)
        {
            if (settings == null || string.IsNullOrWhiteSpace(commandId))
            {
                return;
            }

            settings.RecentCommandIds.RemoveAll(id => string.Equals(id, commandId, StringComparison.OrdinalIgnoreCase));
            settings.RecentCommandIds.Insert(0, commandId);
            Normalize(settings);
        }

        private static void Normalize(UserSettings settings)
        {
            settings.FavoriteCommandIds = Clean(settings.FavoriteCommandIds, int.MaxValue);
            settings.WheelCommandIds = Clean(settings.WheelCommandIds, MaxWheelCommands);
            settings.RecentCommandIds = Clean(settings.RecentCommandIds, MaxRecentCommands);
        }

        private static List<string> Clean(IEnumerable<string> values, int maxCount)
        {
            if (values == null)
            {
                return new List<string>();
            }

            return values
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(maxCount)
                .ToList();
        }
    }
}