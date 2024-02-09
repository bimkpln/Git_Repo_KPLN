using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Общий класс по подготовке к работе с файлами для обмена
    /// </summary>
    public class ExchangeEnvironment
    {
        /// <summary>
        /// Событие, которое возникает при изменении поля для передачи
        /// </summary>
        internal event EventHandler<FieldChangedEventArgs> FieldChanged;

        private protected HashSet<string> _sourceProjectsName = new HashSet<string>(); 
        private protected OpenOptions _openOptions;
        private protected SaveAsOptions _saveAsOptions;

        private static RevitDocExchangestDbService _revitDocExchangestDbService;
        private static UserDbService _userDbService;
        private static RevitDialogDbService _dialogDbService;
        private static ProjectDbService _projectDbService;
        
        private DBUser _dBUser;
        private string _currentDocName;

        internal protected static RevitDocExchangestDbService CurrentRevitDocExchangestDbService
        {
            get
            {
                if (_revitDocExchangestDbService == null)
                    _revitDocExchangestDbService = (RevitDocExchangestDbService)new CreatorRevitDocExchangesDbService().CreateService();
                return _revitDocExchangestDbService;
            }
        }

        internal protected static UserDbService CurrentUserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
                return _userDbService;
            }
        }

        internal protected static RevitDialogDbService CurrentRevitDialogDbService
        {
            get
            {
                if (_dialogDbService == null)
                    _dialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
                return _dialogDbService;
            }
        }

        internal static ProjectDbService CurrentProjectDbService
        {
            get
            {
                if (_projectDbService == null)
                    _projectDbService = (ProjectDbService)new CreatorProjectDbService().CreateService();
                return _projectDbService;
            }
        }

        internal protected static UIControlledApplication RevitUIControlledApp { get; set; }

        internal protected static NLog.Logger Logger { get; set; }

        internal protected static string RevitVersion { get; set; }

        /// <summary>
        /// Счтетчик успешно отработанных процессов
        /// </summary>
        internal protected int CountProcessedDocs { get; set; } = 0;

        /// <summary>
        /// Счтетчик файлов для обработки
        /// </summary>
        internal protected int CountSourceDocs { get; set; } = 0;

        /// <summary>
        /// Список диалогов из БД
        /// </summary>
        internal protected DBRevitDialog[] DBRevitDialogs
        {
            get => CurrentRevitDialogDbService.GetDBRevitDialogs().ToArray();
        }

        /// <summary>
        /// Ссылка на массив документов для обмена данными
        /// </summary>
        internal protected DBRevitDocExchanges[] DBRevitDocExchanges
        {
            get => CurrentRevitDocExchangestDbService.GetDBRevitActiveDocExchanges().ToArray();
        }

        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal protected DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                    _dBUser = CurrentUserDbService.GetCurrentDBUser();

                return _dBUser;
            }
        }

        /// <summary>
        /// Имя документа в текущем процессе обработки
        /// </summary>
        internal protected string CurrentDocName
        {
            get => _currentDocName;
            set
            {
                if (_currentDocName != value)
                {
                    _currentDocName = value;
                    OnFieldChanged(new FieldChangedEventArgs(value));
                }
            }
        }

        internal static void SetStaticEnvironment(UIControlledApplication application, NLog.Logger logger, string revitVersion)
        {
            RevitUIControlledApp = application;
            Logger = logger;
            RevitVersion = revitVersion;
        }


        /// <summary>
        /// Обработчик события
        /// </summary>
        /// <param name="e"></param>
        private protected void OnFieldChanged(FieldChangedEventArgs e)
        {
            FieldChanged?.Invoke(this, e);
        }

        /// <summary>
        /// Подготовка опций к открытию и сохранению
        /// </summary>
        private protected void SetOptions()
        {
            _openOptions = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets };
            _openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

            _saveAsOptions = new SaveAsOptions() { OverwriteExistingFile = true };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions() { SaveAsCentral = true };
            _saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);
        }

        /// <summary>
        /// Отправка результата пользователю в месенджер
        /// </summary>
        private protected virtual void SendResultMsg(string moduleName)
        {
            if (CountProcessedDocs < CountSourceDocs)
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано с ошибками.\n" +
                        $"Метрик производительности: Выгружено {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта/-ов: {string.Join(", ", _sourceProjectsName)}\n" +
                        $"Ошибки: См. файл логов у пользователя {CurrentDBUser.Surname} {CurrentDBUser.Name}.\n" +
                        $"Путь к логам у пользователя: C:\\TEMP\\KPLN_Logs\\2023");
            }
            else
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано без ошибок.\n" +
                        $"Метрик производительности: Обработано {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта/-ов: {string.Join(", ", _sourceProjectsName)}");
            }
        }

        /// <summary>
        /// Подготовка путей к открытию в Revit
        /// </summary>
        /// <param name="pathFrom"></param>
        /// <returns></returns>
        private protected List<string> PreparePathesToOpen(string pathFrom)
        {
            List<string> fileFromPathes = new List<string>();
            
            // Проверяю, что это файл, если нет - то нужно забрать ВСЕ файлы из папки
            if (System.IO.File.Exists(pathFrom))
                fileFromPathes.Add(pathFrom);
            // Проверяю, что это папка, если нет - то нужно забрать ВСЕ файлы из ревит-сервера
            else if (Directory.Exists(pathFrom))
                fileFromPathes = Directory.GetFiles(pathFrom, "*" + ".rvt").ToList<string>();
            // Обработка Revit-Server, чтобы забрать файл или ВСЕ файлы из папки
            // https://www.nuget.org/packages/RevitServerAPILib
            else
            {
                string[] pathParts = pathFrom.Split('\\');

                string rsHostName = pathParts[2];
                if (rsHostName == null)
                {
                    Logger.Error($"Ошибка заполнения пути для копирования с Revit-Server: ({pathFrom}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'");
                    return null;
                }
                try
                {
                    RevitServer server = new RevitServer(rsHostName, int.Parse(RevitVersion));
                    FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathParts.Length - 3));
                    foreach (var model in folderContents.Models)
                    {
                        fileFromPathes.Add($"RSN:{pathFrom}{model.Name}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Ошибка открытия Revit-Server ({pathFrom}):\n{ex.Message}");
                    return null;
                }
            }

            if (fileFromPathes.Count == 0)
            {
                Logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return null;
            }

            return fileFromPathes;
        }
    }
}
