using Autodesk.Revit.UI;
using KPLN_Loader.Core;
using KPLN_Loader.Core.Entities;
using KPLN_Loader.Forms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для подготовки окружения для копируемых плагинов (директории, папки и т.п.)
    /// </summary>
    public sealed class EnvironmentService
    {
        private static Bitrix_Config[] _bitrixConfigs;
        private static DB_Config[] _databaseConfigs;

        ///<summary>
        ///Путь до локальной папки пользователя
        ///</summary>
        private readonly DirectoryInfo _userLocation = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"AppData\Local"));
        ///<summary>
        ///Папка, в которую будут скопированы локально файлы для каждой версии Revit
        ///</summary>
        private readonly DirectoryInfo _applicationLocation;
        ///<summary>
        ///Путь, по которому будет создана папка сесии
        ///</summary>
        private readonly DirectoryInfo _sessionLocation;
        ///<summary>
        ///Версия Revit
        ///</summary>
        private readonly string _revitVersion;
        private readonly Logger _logger;
        private readonly LoaderStatusForm _loaderStatusForm;

        internal EnvironmentService(Logger logger, LoaderStatusForm loaderStatusForm, string revitVersion, string diteTime)
        {
            _logger = logger;
            _loaderStatusForm = loaderStatusForm;
            _revitVersion = revitVersion;
            _applicationLocation = new DirectoryInfo(Path.Combine(_userLocation.FullName, "KPLN_Loader"));
            _sessionLocation = new DirectoryInfo(Path.Combine(_applicationLocation.FullName, $"{_revitVersion}_{diteTime}"));
            ModulesLocation = new DirectoryInfo(Path.Combine(_sessionLocation.FullName, "Modules"));
        }

        /// <summary>
        /// Имя основной БД для работы из конфигураций
        /// </summary>
        public static string DatabaseConfigs_LoaderMainDB { get; } = "Loader_MainDB";

        /// <summary>
        /// Коллекция десерилизованныйх данных по БД
        /// </summary>
        public static DB_Config[] DatabaseConfigs
        {
            get
            {
                if (_databaseConfigs == null)
                    _databaseConfigs = GetDBCongigs();

                return _databaseConfigs;
            }
        }

        /// <summary>
        /// Имя вебхука для работы с BitrixAPI из конфигураций
        /// </summary>
        internal static string BitrixConfigs_MainWebHookName { get; } = "MainWebHook";

        /// <summary>
        /// Коллекция десерилизованныйх данных по настройкам Bitrix
        /// </summary>
        internal static Bitrix_Config[] BitrixConfigs
        {
            get
            {
                if (_bitrixConfigs == null)
                    _bitrixConfigs = GetBitrixCongigs();

                return _bitrixConfigs;
            }
        }

        ///<summary>
        ///Путь, по которому будет создана папка со скопироваными модулями
        ///</summary>
        internal DirectoryInfo ModulesLocation { get; }

        /// <summary>
        /// Получить значение Id пользователя из Битрикс24 КПЛН.
        /// </summary>
        /// <param name="name">Имя пользователя</param>
        /// <param name="surname">Фамилия пользователя</param>
        internal async Task<int> GetUserBitrixId_ByNameAndSurname(string name, string surname)
        {
            int id = -1;
            string webHookUrl = BitrixConfigs.FirstOrDefault(d => d.Name == BitrixConfigs_MainWebHookName).URL;
            using (HttpClient client = new HttpClient())
            {
                // Выполнение GET - запроса к странице
                HttpResponseMessage response = await client.GetAsync($"{webHookUrl}/user.search.json?NAME={name}&LAST_NAME={surname}");
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrEmpty(content))
                        throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                    dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                    dynamic responseResult = dynDeserilazeData.result;
                    // Возвращаю дефолтное значение 
                    if (responseResult.Count == 0)
                        throw new Exception("\n[KPLN]: Произошла ошибка поиска пользователя по Id. Сокрее всего была ошибка при вводе ФИО. Запусти Revit заново, если ошибка не пропадёт - сввяжись с BIM-отделом\n\n");

                    if (!int.TryParse(responseResult[0].ID.ToString(), out id))
                        throw new Exception("\n[KPLN]: Не удалось привести значение ID к int\n\n");
                }
            }

            if (id == -1)
                throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД - не удалось получить id-пользователя Bitrix\n\n");

            return id;
        }

        /// <summary>
        /// Проверка наличия файла конфигурации и необходимых БД
        /// </summary>
        internal EnvironmentService ConfigFileChecker()
        {
            string userErrorMsg = string.Empty;
            if (File.Exists(Application.MainConfigPath))
            {
                StringBuilder stringBuilder = new StringBuilder();

                // Проверка наличия БД
                foreach (DB_Config db in DatabaseConfigs)
                {
                    string fullPath = db.Path;
                    if (!File.Exists(fullPath))
                        stringBuilder.Append($"Отсутствует файл: {fullPath}\r\n");
                }

                // Проверка инициализации вебхуков
                foreach (Bitrix_Config bitr in BitrixConfigs)
                {
                    if (string.IsNullOrEmpty(bitr.Name) || string.IsNullOrEmpty(bitr.URL))
                        stringBuilder.Append($"Не все вебхуки активированы. Отправь разработчику\r\n");
                }

                if (stringBuilder.Length > 0)
                    userErrorMsg = stringBuilder.ToString().TrimEnd();
            }
            else
                userErrorMsg = $"Отсутствует файл конфигураций баз данных: {Application.MainConfigPath}";

            if (!string.IsNullOrEmpty(userErrorMsg))
            {
                TaskDialog td = new TaskDialog("ОШИБКА")
                {
                    MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                    MainInstruction = userErrorMsg,
                    MainContent = "Проверь, подключен ли диск с помощью проводника, если нет - обратись к системному администратору"
                };
                td.Show();

                _logger.Error(userErrorMsg);
                _loaderStatusForm.Dispatcher.Invoke(() => _loaderStatusForm.Close());
                throw new Exception(userErrorMsg);
            }

            return this;
        }

        /// <summary>
        /// Подготовка и очистка старых директорий для копирования
        /// </summary>
        /// <returns></returns>
        internal EnvironmentService PreparingAndCliningDirectories()
        {
            DirectoryInfo appDirInfo = Directory.CreateDirectory(_applicationLocation.FullName);
            int delDirCount = ClearDirectory(appDirInfo.FullName);

            _logger.Info($"Директории успешно очищены от неиспользуемых папок! Удалено {delDirCount} корневых папок");
            Directory.CreateDirectory(ModulesLocation.FullName);

            return this;
        }

        /// <summary>
        /// Копирование модуля
        /// </summary>
        /// <param name="userModule">Модуль для копирования</param>
        /// <returns>DirectoryInfo скопированного модуля</returns>
        internal DirectoryInfo CopyModule(Module userModule, bool isDebugModule)
        {
            DirectoryInfo trueDirInfo = null;
            string targetDir = Path.Combine(ModulesLocation.FullName, userModule.Name);

            DirectoryInfo moduleDirInfo = new DirectoryInfo(userModule.Path);
            if (moduleDirInfo.Exists)
            {
                DirectoryInfo[] currentDirs = moduleDirInfo.GetDirectories();
                // Для дебаг статуса и библиотек выбрана спец. структура директорий - разделена по версиям ревит, только все спрятано в папку Debug
                if (isDebugModule && userModule.IsLibraryModule)
                {
                    IEnumerable<DirectoryInfo> debugDirs = currentDirs.Where(dir => dir.Name.Equals("Debug"));
                    if (debugDirs.Any())
                        trueDirInfo = new DirectoryInfo(Path.Combine(debugDirs.FirstOrDefault().FullName, _revitVersion));
                }
                else
                {
                    IEnumerable<DirectoryInfo> revitVersionDirs = currentDirs.Where(dir => dir.Name.Equals(_revitVersion));
                    // Забираю по папкам для версий Ревит
                    if (revitVersionDirs.Any())
                        trueDirInfo = revitVersionDirs.FirstOrDefault();
                }
            }

            if (trueDirInfo == null || !trueDirInfo.Exists)
            {
                _logger.Error($"Ошибка при проверке наличия модуля {userModule.Name} - путь в БД указан не верно: {userModule.Path}");
                return null;
            }

            CopyDirectory(trueDirInfo.FullName, targetDir);

            return new DirectoryInfo(targetDir);
        }

        /// <summary>
        /// Получить коллекцию Bitrix_Config
        /// </summary>
        /// <returns></returns>
        private static Bitrix_Config[] GetBitrixCongigs()
        {
            string jsonConfig = File.ReadAllText(Application.MainConfigPath);
            JObject root = JObject.Parse(jsonConfig);

            var bitrixSection = root["BitrixConfig"]?["WEBHooks"];
            var bitrixList = new List<Bitrix_Config>();

            if (bitrixSection != null)
            {
                var bitrixObj = bitrixSection.ToObject<Bitrix_Config>();
                if (bitrixObj != null)
                    bitrixList.Add(bitrixObj);
            }

            return bitrixList.ToArray();
        }

        /// <summary>
        /// Получить коллекцию DB_Config
        /// </summary>
        /// <returns></returns>
        private static DB_Config[] GetDBCongigs()
        {
            string jsonConfig = File.ReadAllText(Application.MainConfigPath);
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

        /// <summary>
        /// Очистка (по возможности) указанной директории
        /// </summary>
        /// <param name="directoryPath">Путь для удаления</param>
        private int ClearDirectory(string directoryPath)
        {
            int count = 0;
            try
            {
                foreach (var dirPath in Directory.GetDirectories(directoryPath))
                {
                    if (!IsDirLocked(dirPath))
                    {
                        // Если удалять БЕЗ проверки на занятый файл, то данный метод удалит ВСЁ, что сможет
                        Directory.Delete(dirPath, true);
                        count++;
                    }
                }
            }
            // Что-то пошло не так - логируем
            catch (Exception ex)
            {
                _logger.Error($"При попытке удаления каталога произошла ошибка '{directoryPath}': {ex.Message}");
            }

            return count;
        }

        /// <summary>
        /// Проверка папки на наличие хотя бы 1го занятого.
        /// </summary>
        /// <param name="directoryPath">Папка для анализа</param>
        /// <returns></returns>
        private bool IsDirLocked(string directoryPath)
        {
            string[] filePathes = Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories);

            foreach (string filePath in filePathes)
            {
                try
                {
                    using (var stream = new FileInfo(filePath).Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        stream.Close();
                    }
                }
                // Файл занят
                catch (IOException)
                {
                    return true;
                }

            }

            // Все файлы свободны
            return false;
        }

        /// <summary>
        /// Копирование файлов в указанную директорию
        /// </summary>
        /// <param name="sourceDir">Путь откуда</param>
        /// <param name="destDir">Путь куда</param>
        private void CopyDirectory(string sourceDir, string destDir)
        {
            if (!Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            foreach (string sourcePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(sourcePath);
                string targetPath = Path.Combine(destDir, fileName);
                File.Copy(sourcePath, targetPath, false);
            }

            foreach (string sourceSubDir in Directory.GetDirectories(sourceDir))
            {
                string dirName = Path.GetFileName(sourceSubDir);
                string targetSubDir = Path.Combine(destDir, dirName);
                CopyDirectory(sourceSubDir, targetSubDir);
            }
        }
    }
}
