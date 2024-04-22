using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.IO;

namespace KPLN_BIMTools_Ribbon.Core.SQLite
{
    internal static class DBEnvironment
    {
        private const string _mainFolderPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\DocExchange_Settings";

        /// <summary>
        /// Генерация пути БД для хранения данных по ReportInstance
        /// </summary>
        internal static FileInfo GenerateNewPath(DBProject dBProject, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            int index = 1;
            // Проверка на наличие файла с таким именем.
            while (File.Exists(Path.Combine(_mainFolderPath, $"{revitDocExchangeEnum}_PRG_{dBProject.Id}_{index}.db")))
                index++;

            return new FileInfo(Path.Combine(_mainFolderPath, $"{revitDocExchangeEnum}_PRG_{dBProject.Id}_{index}.db"));
        }
    }
}
