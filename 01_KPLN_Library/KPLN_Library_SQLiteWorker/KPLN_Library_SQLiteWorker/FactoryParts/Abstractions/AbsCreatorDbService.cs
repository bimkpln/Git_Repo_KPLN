using KPLN_Library_SQLiteWorker.Core;
using KPLN_Library_SQLiteWorker.FactoryParts.Abstractions;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_Library_SQLiteWorker.Abstractions
{
    /// <summary>
    /// Абстрактный создатель сервисов работы с БД
    /// </summary>
    public abstract class AbsCreatorDbService
    {
        private protected static readonly string _sqlMainConfigPath = KPLN_Loader.Application.SQLMainConfigPath;
        private List<DatabasePaths> _databasesPaths;

        /// <summary>
        /// Создать инстанс необходимого сервиса
        /// </summary>
        /// <returns></returns>
        public abstract AbsDbService CreateService();

        /// <summary>
        /// Проверка наличия файла конфига и самих файлов БД
        /// </summary>
        /// <exception cref="Exception">Если чего-то нет - выброс ошибки</exception>
        private protected void SQLFilesExistChecker()
        {
            string userErrorMsg = string.Empty;
            if (File.Exists(_sqlMainConfigPath))
            {
                StringBuilder stringBuilder = new StringBuilder();

                string jsonConfig = File.ReadAllText(_sqlMainConfigPath);
                _databasesPaths = JsonConvert.DeserializeObject<List<DatabasePaths>>(jsonConfig);
                foreach (DatabasePaths db in _databasesPaths)
                {
                    string fullPath = db.Path;
                    if (!File.Exists(fullPath))
                    {
                        stringBuilder.Append($"Отсутствует файл: {fullPath}\r\n");
                    }
                }

                if (stringBuilder.Length > 0)
                    userErrorMsg = stringBuilder.ToString().TrimEnd();
            }
            else
            {
                userErrorMsg = $"Отсутствует файл: {_sqlMainConfigPath}";
            }

            if (!string.IsNullOrEmpty(userErrorMsg))
                throw new Worker_Error(userErrorMsg);
        }

        /// <summary>
        /// Подготовка пути для подключения к БД
        /// </summary>
        private protected string CreateConnectionString(string currentPath) => $"Data Source={_databasesPaths.FirstOrDefault(d => d.Name.Contains(currentPath)).Path}; Version=3;";
    }
}
