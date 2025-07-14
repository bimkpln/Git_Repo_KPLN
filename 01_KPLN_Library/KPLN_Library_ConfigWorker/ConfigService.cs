using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
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
        // Хранится на диске C:\\
        Local,
        // Хранится на диске Z:\\
        Shared,
    }

    /// <summary>
    /// Сервис, обсулживающий процесс сохранения файлов-конфигураций
    /// </summary>
    public static class ConfigService
    {
        private static readonly ProjectDbService _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
        private static readonly string _localConfigFolder = $"{Path.GetPathRoot(Environment.SystemDirectory)}KPLN_Temp";

        /// <summary>
        /// Кэш сущности в память
        /// </summary>
        public static object MemoryConfigData { get; private set; }

        /// <summary>
        /// Запись данных в файл конфига
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="data">Данные для записи (один объект или коллекция объектов)</param>
        /// <param name="configType">Тип конфигурации для сохранения</param>
        /// <param name="configName">Имя конфигурации</param>
        public static void SaveConfig<T>(Document doc, ConfigType configType, object data, string configName = "") where T : IJsonSerializable
        {
            string jsonEntity;

            if (data is IEnumerable<T> collection)
                jsonEntity = JsonConvert.SerializeObject(collection.Select(ent => ent.ToJson()), Formatting.Indented);
            else if (data is T singleEntity)
                jsonEntity = JsonConvert.SerializeObject(singleEntity.ToJson(), Formatting.Indented);
            else
                throw new ArgumentException("Ошибка передаваемой сущности. Должна быть: IJsonSerializable.");

            if (configType == ConfigType.Memory)
                MemoryConfigData = data;
            else
            {
                string configPath = CreateConfigPath(doc, configName, configType);

                if (!new FileInfo(configPath).Exists)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(configPath));
                    FileStream fileStream = File.Create(configPath);
                    fileStream.Dispose();
                }

                using (StreamWriter streamWriter = new StreamWriter(configPath))
                {
                    streamWriter.Write(jsonEntity);
                }
            }
        }

        /// <summary>
        /// Десереилизация конфига с получением результата
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="configType">Тип конфигурации для сохранения</param>
        /// <param name="configName">Имя конфигурации</param>
        public static object ReadConfigFile<T>(Document doc, ConfigType configType, string configName="")
        {
            string json = string.Empty;

            switch (configType)
            {
                case ConfigType.Memory:
                    return MemoryConfigData;
                case ConfigType.Local:
                    goto case ConfigType.Shared;
                case ConfigType.Shared:
                    string configPath = CreateConfigPath(doc, configName, configType);
                    if (new FileInfo(configPath).Exists)
                    {
                        using (StreamReader streamReader = new StreamReader(configPath))
                        {
                            json = streamReader.ReadToEnd();
                        }
                    }
                    try
                    {
                        if (typeof(T).IsGenericType && typeof(T).GetGenericTypeDefinition() == typeof(IEnumerable<>))
                            return JsonConvert.DeserializeObject<IEnumerable<T>>(json);
                        else
                            return JsonConvert.DeserializeObject<T>(json);
                    }
                    catch (JsonException ex)
                    {
                        throw new InvalidOperationException($"Failed to deserialize config file: {ex.Message}", ex);
                    }
            }

            return null;
        }

        /// <summary>
        /// Генерация пути к файлу-конфигурации
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="configName">Имя конфигурации</param>
        /// <param name="configType">Тип конфигурации для сохранения</param>
        private static string CreateConfigPath(Document doc, string configName, ConfigType configType)
        {
            ModelPath docModelPath = doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);

            if (string.IsNullOrEmpty(configName))
                throw new Exception("Системная ошибка имени конфигурации. Обратись к разработчику!");

            switch (configType)
            {
                // Сохраняется в системной папке пользователя
                case ConfigType.Local:
                    return $"{_localConfigFolder}\\KPLN_Config\\{configName}.json";
                // Сохраняется в папке на диске Z
                case ConfigType.Shared:
                    string resultPath;
                    if (strDocModelPath.Contains("RSN:"))
                    {
                        DBProject dBProject = _projectDbService.GetDBProject_ByRevitDocFileName(strDocModelPath);
                        resultPath = $"Z:\\KPLN_Temp\\KPLN_Config\\{dBProject.Code}\\{configName}.json";
                    }
                    else
                    {
                        string trimmedPath = strDocModelPath.Trim($"{doc.Title}.rvt".ToArray());
                        resultPath = trimmedPath + $"KPLN_Config\\{configName}.json";
                    }

                    return resultPath;
            }

            throw new Exception("Системная ошибка выбора типа хранения конфигурации. Обратись к разработчику!");
        }
    }
}
