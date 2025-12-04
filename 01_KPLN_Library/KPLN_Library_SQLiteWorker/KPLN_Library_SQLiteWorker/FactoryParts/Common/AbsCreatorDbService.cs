using KPLN_Library_SQLiteWorker.Core;
using KPLN_Loader.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_Library_SQLiteWorker.FactoryParts.Common
{
    /// <summary>
    /// Абстрактный создатель сервисов работы с БД
    /// </summary>
    public abstract class AbsCreatorDbService
    {
        /// <summary>
        /// Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
        /// </summary>
        private protected static readonly string _mainConfigPath = @"Z:\Отдел BIM\03_Скрипты\08_Базы данных\KPLN_Loader_Config.json";
        /// <summary>
        /// Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
        /// </summary>
        private protected static readonly string _mainDBName = "Loader_MainDB";
        private DB_Config[] _databaseConfigs;

        private DB_Config[] DatabaseConfigs
        {
            get
            {
                if (_databaseConfigs == null)
                    _databaseConfigs = GetDBCongigs();

                return _databaseConfigs;
            }
        }

        /// <summary>
        /// Создать инстанс необходимого сервиса
        /// </summary>
        /// <returns></returns>
        public abstract DbService CreateService();

        /// <summary>
        /// Проверка наличия файла конфига и самих файлов БД, и получение конфигов
        /// </summary>
        /// <exception cref="Exception">Если чего-то нет - выброс ошибки</exception>
        private protected void SQLFilesExistCheckerAndDBDataSetter()
        {
            string userErrorMsg = string.Empty;
            if (File.Exists(_mainConfigPath))
            {
                StringBuilder stringBuilder = new StringBuilder();

                foreach (DB_Config db in DatabaseConfigs)
                {
                    string fullPath = db.Path;
                    if (!File.Exists(fullPath))
                        stringBuilder.Append($"Отсутствует файл: {fullPath}\r\n");
                }

                if (stringBuilder.Length > 0)
                    userErrorMsg = stringBuilder.ToString().TrimEnd();
            }
            else
            {
                userErrorMsg = $"Отсутствует файл: {_mainConfigPath}";
            }

            if (!string.IsNullOrEmpty(userErrorMsg))
                throw new Worker_Error(userErrorMsg);
        }

        /// <summary>
        /// Подготовка пути для подключения к БД
        /// </summary>
        private protected string CreateConnectionString() => $"Data Source={DatabaseConfigs.FirstOrDefault(d => d.Name.Contains(_mainDBName)).Path}; Version=3;";

        /// <summary>
        /// Получить конфиги по БД. Лучше заменить на ссылку на KPLN_Loader. Пока захаркодил, т.к. KPLN_Loader нужно всем переустановить (21.10.2025)
        /// </summary>
        /// <returns></returns>
        private DB_Config[] GetDBCongigs()
        {
            string jsonConfig = File.ReadAllText(_mainConfigPath);
            JObject root = JObject.Parse(jsonConfig);

            var dbSection = root["DatabaseConfig"]?["DBConnections"] as JObject;
            var dbList = new List<DB_Config>();

            if (dbSection != null)
            {
                foreach (var prop in dbSection.Properties())
                {
                    var dbObj = prop.Value.ToObject<DB_Config>();
                    if (dbObj != null)
                        dbList.Add(dbObj);
                }
            }

            return dbList.ToArray();
        }
    }
}
