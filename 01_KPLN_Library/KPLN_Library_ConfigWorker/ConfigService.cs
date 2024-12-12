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
    /// <summary>
    /// Сервис, обсулживающий процесс сохранения файлов-конфигураций
    /// </summary>
    public static class ConfigService
    {
        private static readonly ProjectDbService _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
        private static readonly string _localConfigFolder = $"{Path.GetPathRoot(Environment.SystemDirectory)}KPLN_Temp";

        /// <summary>
        /// Запись данных в файл конфига
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="configName">Имя конфигурации</param>
        /// <param name="data">Данные для записи (один объект или коллекция объектов)</param>
        /// <param name="isLocalConfig">Локальный конфиг? Если да, то сохраняется в системной папке пользователя</param>
        public static void SaveConfig<T>(Document doc, string configName, object data, bool isLocalConfig) where T : IJsonSerializable
        {
            string configPath = CreateConfigPath(doc, configName, isLocalConfig);

            if (!new FileInfo(configPath).Exists)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(configPath));
                FileStream fileStream = File.Create(configPath);
                fileStream.Dispose();
            }

            using (StreamWriter streamWriter = new StreamWriter(configPath))
            {
                string jsonEntity;

                if (data is IEnumerable<T> collection)
                    jsonEntity = JsonConvert.SerializeObject(collection.Select(ent => ent.ToJson()), Formatting.Indented);
                else if (data is T singleEntity)
                    jsonEntity = JsonConvert.SerializeObject(singleEntity.ToJson(), Formatting.Indented);
                else
                    throw new ArgumentException("Ошибка передаваемой сущности. Должна быть: IJsonSerializable.");

                streamWriter.Write(jsonEntity);
            }
        }

        /// <summary>
        /// Десереилизация конфига с получением результата
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="configName">Имя конфигурации</param>
        /// <param name="isLocalConfig">Локальный конфиг? Если да, то сохраняется в системной папке пользователя</param>
        public static object ReadConfigFile<T>(Document doc, string configName, bool isLocalConfig)
        {
            string configPath = CreateConfigPath(doc, configName, isLocalConfig);
            if (new FileInfo(configPath).Exists)
            {
                using (StreamReader streamReader = new StreamReader(configPath))
                {
                    string json = streamReader.ReadToEnd();

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
            }

            return default;
        }

        /// <summary>
        /// Генерация пути к файлу-конфигурации
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="configName">Имя конфигурации</param>
        /// <param name="isLocalConfig">Локальный конфиг? Если да, то сохраняется в системной папке пользователя</param>
        private static string CreateConfigPath(Document doc, string configName, bool isLocalConfig)
        {
            ModelPath docModelPath = doc.GetWorksharingCentralModelPath() ?? throw new Exception("Работает только с моделями из хранилища");
            string strDocModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(docModelPath);

            if (isLocalConfig)
                return $"{_localConfigFolder}\\KPLN_Config\\{configName}.json";

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
    }
}
