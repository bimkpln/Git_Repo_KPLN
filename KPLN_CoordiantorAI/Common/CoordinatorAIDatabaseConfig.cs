using System;
using System.IO;

namespace KPLN_CoordiantorAI.Common
{
    public static class CoordinatorAiDatabaseConfig
    {
        public const string DefaultDatabaseFolder =
            "Z:\\\u041E\u0442\u0434\u0435\u043B BIM\\03_\u0421\u043A\u0440\u0438\u043F\u0442\u044B\\08_\u0411\u0430\u0437\u044B \u0434\u0430\u043D\u043D\u044B\u0445";

        public const string DatabaseName = "KPLN_CoordinatorAI";

        public const string DatabaseFileName = DatabaseName + ".db";

        private static readonly string DatabasePathSettingsFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "KPLN",
            "CoordinatorAI");

        private static readonly string DatabasePathSettingsFile = Path.Combine(
            DatabasePathSettingsFolder,
            "database-path.txt");

        private static string _databaseFilePath = LoadDatabaseFilePath();

        public static string DatabaseFolder
        {
            get { return Path.GetDirectoryName(DatabaseFilePath); }
        }

        public static string DatabaseFilePath
        {
            get { return _databaseFilePath; }
        }

        public static void SetDatabaseFilePath(string databaseFilePath)
        {
            string normalizedPath = (databaseFilePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedPath))
                normalizedPath = GetDefaultDatabaseFilePath();

            _databaseFilePath = normalizedPath;

            if (!Directory.Exists(DatabasePathSettingsFolder))
                Directory.CreateDirectory(DatabasePathSettingsFolder);

            File.WriteAllText(DatabasePathSettingsFile, _databaseFilePath);
        }

        private static string LoadDatabaseFilePath()
        {
            if (File.Exists(DatabasePathSettingsFile))
            {
                string savedPath = File.ReadAllText(DatabasePathSettingsFile).Trim();
                if (!string.IsNullOrWhiteSpace(savedPath))
                    return savedPath;
            }

            return GetDefaultDatabaseFilePath();
        }

        private static string GetDefaultDatabaseFilePath()
        {
            return Path.Combine(DefaultDatabaseFolder, DatabaseFileName);
        }
    }
}
