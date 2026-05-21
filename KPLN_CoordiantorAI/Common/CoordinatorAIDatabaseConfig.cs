using System.IO;

namespace KPLN_CoordiantorAI.Common
{
    public static class CoordinatorAiDatabaseConfig
    {
        public const string DatabaseFolder = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных";

        public const string DatabaseName = "KPLN_CoordinatorAI";

        public const string DatabaseFileName = DatabaseName + ".db";

        public static string DatabaseFilePath
        {
            get { return Path.Combine(DatabaseFolder, DatabaseFileName); }
        }
    }
}
