using Autodesk.Revit.UI;
using KPLN_Loader.Core;
using KPLN_Loader.Core.SQLiteData;
using Newtonsoft.Json;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace KPLN_Loader.Services
{
    /// <summary>
    /// Сервис для подготовки окружения для копируемых плагинов (директории, папки и т.п.)
    /// </summary>
    internal class EnvironmentService
    {
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
        ///Путь, по которому будет создана папка со скопироваными модулями
        ///</summary>
        private readonly DirectoryInfo _modulesLocation;
        ///<summary>
        ///Версия Revit
        ///</summary>
        private readonly string _revitVersion;
        private readonly Logger _logger;

        public EnvironmentService(Logger logger, string revitVersion, string diteTime)
        {
            _logger = logger;
            _revitVersion = revitVersion;
            _applicationLocation = new DirectoryInfo(Path.Combine(_userLocation.FullName, "KPLN_Loader"));
            _sessionLocation = new DirectoryInfo(Path.Combine(_applicationLocation.FullName, $"{_revitVersion}_{diteTime}"));
            _modulesLocation = new DirectoryInfo(Path.Combine(_sessionLocation.FullName, "Modules"));
        }

        /// <summary>
        /// Коллекция десерилизованныйх данных по БД
        /// </summary>
        public List<DatabasePaths> DatabasesPaths { get; private set; }

        ///<summary>
        ///Путь, по которому будет создана папка со скопироваными модулями
        ///</summary>
        public DirectoryInfo ModulesLocation 
        {
            get { return _modulesLocation; }
        }

        /// <summary>
        /// Подготовка и очистка старых директорий для копирования
        /// </summary>
        /// <returns></returns>
        public void PreparingAndCliningDirectories()
        {
            DirectoryInfo appDirInfo = Directory.CreateDirectory(_applicationLocation.FullName);
            int delDirCount = 0;
            foreach (DirectoryInfo subLoc in appDirInfo.GetDirectories())
                delDirCount = ClearDirectory(subLoc.FullName);
            
            _logger.Info($"Директории успешно очищены от неиспользуемых папок! Удалено {delDirCount} корневых папок");
            Directory.CreateDirectory(_modulesLocation.FullName);
        }

        /// <summary>
        /// Проверка наличия необходимых БД
        /// </summary>
        public void SQLFilesExistChecker(string sqlConfigPath)
        {
            string userErrorMsg = string.Empty;
            if (File.Exists(sqlConfigPath))
            {
                StringBuilder stringBuilder = new StringBuilder();

                string jsonConfig = File.ReadAllText(sqlConfigPath);
                DatabasesPaths = JsonConvert.DeserializeObject<List<DatabasePaths>>(jsonConfig);

                foreach (DatabasePaths db in DatabasesPaths)
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
                userErrorMsg = $"Отсутствует файл: {sqlConfigPath}";
            }

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
                throw new Exception(userErrorMsg);
            }
        }

        /// <summary>
        /// Копирование модуля
        /// </summary>
        /// <param name="userModule">Модуль для копирования</param>
        /// <returns>DirectoryInfo скопированного модуля</returns>
        public DirectoryInfo CopyModule(Module userModule)
        {
            string targetDir = Path.Combine(_modulesLocation.FullName, userModule.Name);
            DirectoryInfo moduleRevitVersionDirInfo = new DirectoryInfo(Path.Combine(userModule.Path, _revitVersion));
            DirectoryInfo moduleDirInfo = new DirectoryInfo(userModule.Path);
            if (moduleRevitVersionDirInfo.Exists)
            {
                CopyDirectory(moduleRevitVersionDirInfo.FullName, targetDir);
            }
            else if (moduleDirInfo.Exists)
            {
                CopyDirectory(moduleDirInfo.FullName, targetDir);
            }
            else
                _logger.Error($"Ошибка при проверке наличия модуля {userModule.Name} - путь в БД указан не верно: {userModule.Path}");

            return new DirectoryInfo(targetDir);
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
                if (Directory.Exists(directoryPath))
                {
                    Directory.Delete(directoryPath, true);
                    count++;
                }
            }
            // Отлов занятых папок - их не удаялем
            catch (UnauthorizedAccessException) { }
            // Что-то пошло не так - логируем
            catch (Exception ex)
            {
                _logger.Error($"При попытке удаления каталога произошла ошибка '{directoryPath}': {ex.Message}");
            }

            return count;
        }

        /// <summary>
        /// Копирование файлов в указанную директорию
        /// </summary>
        /// <param name="sourceDir">Путь откуда</param>
        /// <param name="targetDir">Путь куда</param>
        private void CopyDirectory(string sourceDir, string targetDir)
        {
            if (!Directory.Exists(targetDir))
                Directory.CreateDirectory(targetDir);

            foreach (string sourcePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(sourcePath);
                string targetPath = Path.Combine(targetDir, fileName);
                File.Copy(sourcePath, targetPath, false);
            }

            foreach (string sourceSubDir in Directory.GetDirectories(sourceDir))
            {
                string dirName = Path.GetFileName(sourceSubDir);
                string targetSubDir = Path.Combine(targetDir, dirName);
                CopyDirectory(sourceSubDir, targetSubDir);
            }
        }

    }
}
