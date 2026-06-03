using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_Library_ConfigWorker
{
    public enum ConfigType
    {
        // Храниться в памяти
        Memory,
        // Хранится на диске C:\\ в %appdata%
        Local,
        // Хранится на в корне папки на сервере, или на диске Z:\\ для модлей с RS
        Shared,
    }

    /// <summary>
    /// Сервис, обсулживающий процесс сохранения файлов-конфигураций
    /// </summary>
    public static class ConfigService
    {
        private static readonly string _localConfigFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KPLN_Temp");

        /// <summary>
        /// Кэш сущности в память
        /// </summary>
        public static object MemoryConfigData { get; private set; }

        /// <summary>
        /// Запись данных в файл конфига (без привязки к документу — только Memory / Local)
        /// </summary>
        public static void SaveConfig<T>(ConfigType configType, object data, string configName = "")
            where T : IJsonSerializable
        {
            string json = Serialize<T>(data);

            switch (configType)
            {
                case ConfigType.Memory:
                    SaveMemory(data);
                    break;
                case ConfigType.Local:
                    SaveLocal(configName, json);
                    break;
                case ConfigType.Shared:
                    throw new ArgumentException(
                        "Для Shared-конфига необходимо передать Document. Используй перегрузку с параметром doc.");
            }
        }

        /// <summary>
        /// Запись данных в файл конфига (с привязкой к документу — поддерживает все типы)
        /// </summary>
        public static void SaveConfig<T>(int revitVersion, Document doc, ConfigType configType, object data, string configName = "")
            where T : IJsonSerializable
        {
            string json = Serialize<T>(data);

            switch (configType)
            {
                case ConfigType.Memory:
                    SaveMemory(data);
                    break;
                case ConfigType.Local:
                    SaveLocal(configName, json);
                    break;
                case ConfigType.Shared:
                    SaveShared(revitVersion, doc, configName, json);
                    break;
            }
        }

        /// <summary>
        /// Чтение конфига без привязки к документу — только Memory / Local
        /// </summary>
        public static object ReadConfigFile<T>(ConfigType configType, string configName = "")
        {
            switch (configType)
            {
                case ConfigType.Memory:
                    return ReadFromMemory();
                case ConfigType.Local:
                    return ReadFromLocal<T>(configName);
                case ConfigType.Shared:
                    throw new ArgumentException(
                        "Для Shared-конфига необходимо передать Document. Используй перегрузку с параметром doc.");
            }

            return null;
        }

        /// <summary>
        /// Чтение конфига с привязкой к документу — поддерживает все типы
        /// </summary>
        public static object ReadConfigFile<T>(int revitVersion, Document doc, ConfigType configType, string configName = "")
        {
            switch (configType)
            {
                case ConfigType.Memory:
                    return ReadFromMemory();
                case ConfigType.Local:
                    return ReadFromLocal<T>(configName);
                case ConfigType.Shared:
                    return ReadFromShared<T>(revitVersion, doc, configName);
            }

            return null;
        }

        private static void SaveMemory(object data)
        {
            MemoryConfigData = data;
        }

        private static void SaveLocal(string configName, string json)
        {
            string path = CreateConfigPath_Local(configName);
            WriteToFile(path, json);
        }

        private static void SaveShared(int revitVersion, Document doc, string configName, string json)
        {
            string path = CreateConfigPath(revitVersion, doc, configName, ConfigType.Shared);
            WriteToFile(path, json);
        }

        /// <summary>
        /// Сериализация одного объекта или коллекции в JSON
        /// </summary>
        private static string Serialize<T>(object data) where T : IJsonSerializable
        {
            if (data is IEnumerable<T> collection)
                return JsonConvert.SerializeObject(collection.Select(e => e.ToJson()), Formatting.Indented);

            if (data is T single)
                return JsonConvert.SerializeObject(single.ToJson(), Formatting.Indented);

            throw new ArgumentException("Ошибка передаваемой сущности. Должна быть: IJsonSerializable.");
        }

        /// <summary>
        /// Создаёт файл (и директорию) если нет, и записывает json
        /// </summary>
        private static void WriteToFile(string path, string json)
        {
            if (!new FileInfo(path).Exists)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.Create(path).Dispose();
            }

            using (StreamWriter sw = new StreamWriter(path))
                sw.Write(json);
        }

        private static string CreateConfigPath_Local(string configName) => $"{_localConfigFolder}\\KPLN_Config\\{configName}.json";

        private static string CreateConfigPath(int revitVersion, Document doc, string configName, ConfigType configType)
        {
            if (string.IsNullOrEmpty(configName))
                throw new Exception("Системная ошибка имени конфигурации. Обратись к разработчику!");

            switch (configType)
            {
                case ConfigType.Local:
                    return CreateConfigPath_Local(configName);

                case ConfigType.Shared:
                    ModelPath docModelPath = doc.GetWorksharingCentralModelPath()
                        ?? throw new Exception("Работает только с моделями из хранилища");

                    string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);

                    if (strDocModelPath.Contains("RSN:"))
                    {
                        DBProject dBProject = SQLiteMainService.SQLitePrjServiceInst
                            .GetDBProject_ByRevitDocFileNameANDRVersion(strDocModelPath, revitVersion);
                        return $"Z:\\KPLN_Temp\\KPLN_Config\\{dBProject.Code}\\{configName}.json";
                    }
                    else
                    {
                        string trimmedPath = strDocModelPath.Trim($"{doc.Title}.rvt".ToArray());
                        return $"{trimmedPath}KPLN_Config\\{configName}.json";
                    }
            }

            throw new Exception("Системная ошибка выбора типа хранения конфигурации. Обратись к разработчику!");
        }

        private static object ReadFromMemory()
        {
            return MemoryConfigData;
        }

        private static object ReadFromLocal<T>(string configName)
        {
            string path = CreateConfigPath_Local(configName);
            return Deserialize<T>(ReadFromFile(path));
        }

        private static object ReadFromShared<T>(int revitVersion, Document doc, string configName)
        {
            string path = CreateConfigPath(revitVersion, doc, configName, ConfigType.Shared);
            return Deserialize<T>(ReadFromFile(path));
        }

        /// <summary>
        /// Читает содержимое файла, возвращает пустую строку если файл не существует
        /// </summary>
        private static string ReadFromFile(string path)
        {
            if (!new FileInfo(path).Exists)
                return string.Empty;

            using (StreamReader sr = new StreamReader(path))
                return sr.ReadToEnd();
        }

        /// <summary>
        /// Десериализация JSON в объект или коллекцию
        /// </summary>
        private static object Deserialize<T>(string json)
        {
            try
            {
                if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    return JsonConvert.DeserializeObject<IEnumerable<T>>(json);

                return JsonConvert.DeserializeObject<T>(json);
            }
            catch (JsonException ex)
            {
                throw new InvalidOperationException($"Failed to deserialize config file: {ex.Message}", ex);
            }
        }
    }
}
