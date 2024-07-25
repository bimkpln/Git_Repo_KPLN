using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_BIMTools_Ribbon.Forms;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
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
    /// Общий сервис по подготовке к работе и работе с файлами для обмена
    /// </summary>
    public class ExchangeService
    {
        /// <summary>
        /// Событие, которое возникает при изменении поля для передачи
        /// </summary>
        internal event EventHandler<FieldChangedEventArgs> FieldChanged;

        private protected string _sourceProjectName;
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

        /// <summary>
        /// Установка общих параметров для запуска
        /// </summary>
        internal static void SetStaticEnvironment(UIControlledApplication application, NLog.Logger logger, string revitVersion)
        {
            RevitUIControlledApp = application;
            Logger = logger;
            RevitVersion = revitVersion;
        }

        /// <summary>
        /// Старт сервиса по обмену моделями
        /// </summary>
        private protected bool StartService(UIApplication uiapp, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm();
            bool? dialogResult = selectedProjectForm.ShowDialog();
            if (selectedProjectForm.Status == UIStatus.RunStatus.Run)
            {
                RevitEventWorker revitEventWorker = new RevitEventWorker(this, Logger, DBRevitDialogs);

                // Подписка на события
                RevitUIControlledApp.DialogBoxShowing += revitEventWorker.OnDialogBoxShowing;
                RevitUIControlledApp.ControlledApplication.DocumentOpened += revitEventWorker.OnDocumentOpened;
                RevitUIControlledApp.ControlledApplication.DocumentClosed += revitEventWorker.OnDocumentClosed;
                //RevitUIControlledApp.ControlledApplication.FailuresProcessing += revitEventWorker.OnFailureProcessing;

                // Локальный try, чтобы гарантированно отписаться от событий. Cath - кидает ошибку выше
                try
                {
                    DBProject dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
                    _sourceProjectName = dBProject.Name;

                    ConfigDispatcher configDispatcher = new ConfigDispatcher(Logger, CurrentRevitDocExchangestDbService, dBProject, revitDocExchangeEnum);
                    configDispatcher.ShowDialog();
                    if (configDispatcher.IsRun)
                    {
                        Logger.Info($"Старт экспорта: [{revitDocExchangeEnum}]");

                        foreach (DBRevitDocExchanges currentDocExchange in configDispatcher.SelectedDBExchangeEntities)
                        {
                            SQLiteService sqliteService = new SQLiteService(Logger, currentDocExchange.SettingDBFilePath, revitDocExchangeEnum);
                            IEnumerable<DBConfigEntity> configs = sqliteService.GetConfigItems();
                            foreach (DBConfigEntity config in configs)
                            {
                                List<string> fileFromPathes = PreparePathesToOpen(config.PathFrom);
                                if (fileFromPathes.Count == 0)
                                {
                                    Logger.Error($"Не удалось найти Revit-файлы из папки: {config.PathFrom}");
                                    return false;
                                }

                                // Проверяю, что это папка, если нет - то ревит-сервер
                                if (Directory.Exists(config.PathTo))
                                {
                                    foreach (string fileFromPath in fileFromPathes)
                                    {
                                        CountSourceDocs++;
                                        string newFilePath = ExchangeFile(uiapp.Application, fileFromPath, config);
                                        if (newFilePath != null)
                                        {
                                            CountProcessedDocs++;
                                            Logger.Info($"Файл {newFilePath} успешно сохранен!\n");
                                        }
                                        else
                                            Logger.Error($"Ошибки с файлом {config.Name} описаны выше.\n");
                                    }
                                }
                                // Обрабатываю ревит-сервер
                                else
                                {
                                    foreach (string fileFromPath in fileFromPathes)
                                    {
                                        CountSourceDocs++;
                                        string newFilePath = ExchangeFile(uiapp.Application, fileFromPath, config, "RSN:");
                                        if (newFilePath != null)
                                        {
                                            CountProcessedDocs++;
                                            Logger.Info($"Файл {newFilePath} успешно сохранен!\n");
                                        }
                                    }
                                }
                            }
                        }

                        SendResultMsg($"Плагин экспорта: [{revitDocExchangeEnum}]");
                        Logger.Info($"Файлы успешно экспортированы: [{revitDocExchangeEnum}]");

                        revitEventWorker.Dispose();

                        return true;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    //Отписка от событий
                    RevitUIControlledApp.DialogBoxShowing -= revitEventWorker.OnDialogBoxShowing;
                    RevitUIControlledApp.ControlledApplication.DocumentOpened -= revitEventWorker.OnDocumentOpened;
                    RevitUIControlledApp.ControlledApplication.DocumentClosed -= revitEventWorker.OnDocumentClosed;
                    //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= revitEventWorker.OnFailureProcessing;
                }
                
            }

            return false;
        }

        /// <summary>
        /// Метод обмена файлами
        /// </summary>
        private protected virtual string ExchangeFile(Application app, string fileFromPath, DBConfigEntity configEntity, string rsn = "")
        {
            throw new NotImplementedException("Ошибка реализации структуры! Нужно переопределить метод ExchangeFiles для каджого экспортера");
        }

        /// <summary>
        /// Подготовка опций к открытию
        /// </summary>
        private protected void SetOpenOptions(WorksetConfigurationOption worksetConfigurationOption)
        {
            _openOptions = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets };
            _openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(worksetConfigurationOption));
        }

        /// <summary>
        /// Подготовка опций к открытию с указанием рабочих наборов
        /// </summary>
        private protected void SetOpenOptions(IList<WorksetId> worksetIds)
        {
            _openOptions = new OpenOptions() 
            { 
                DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets 
            };
            WorksetConfiguration openConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
            openConfig.Open(worksetIds);
            _openOptions.SetOpenWorksetsConfiguration(openConfig);
        }

        /// <summary>
        /// Подготовка опций к сохранению
        /// </summary>
        private protected void SetSaveAsOptions()
        {
            _saveAsOptions = new SaveAsOptions() 
            { 
                OverwriteExistingFile = true 
            };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions() 
            { 
                SaveAsCentral = true, 
                OpenWorksetsDefault = SimpleWorksetConfiguration.AskUserToSpecify 
            };
            _saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);
        }


        /// <summary>
        /// Обработчик события
        /// </summary>
        /// <param name="e"></param>
        private void OnFieldChanged(FieldChangedEventArgs e)
        {
            FieldChanged?.Invoke(this, e);
        }

        /// <summary>
        /// Отправка результата пользователю в месенджер
        /// </summary>
        private void SendResultMsg(string moduleName)
        {
            if (CountProcessedDocs < CountSourceDocs)
            {
                BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                    CurrentDBUser,
                    $"Модуль: {moduleName}\n" +
                    $"Статус: Отработано с ошибками.\n" +
                    $"Метрик производительности: Выгружено {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта: {_sourceProjectName}\n" +
                    $"Ошибки: См. файл логов у пользователя {CurrentDBUser.Surname} {CurrentDBUser.Name}.\n" +
                    $"Путь к логам у пользователя: C:\\TEMP\\KPLN_Logs\\{RevitVersion}");
            }
            else
            {
                BitrixMessageSender.SendMsg_ToUser_ByDBUser(
                    CurrentDBUser,
                    $"Модуль: {moduleName}\n" +
                    $"Статус: Отработано без ошибок.\n" +
                    $"Метрик производительности: Обработано {CountProcessedDocs} из {CountSourceDocs} файлов, для проекта: {_sourceProjectName}");
            }
        }

        /// <summary>
        /// Подготовка путей к открытию в Revit
        /// </summary>
        /// <param name="pathFrom"></param>
        /// <returns></returns>
        private List<string> PreparePathesToOpen(string pathFrom)
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
                int pathPartsLenght = pathParts.Length;
                if (rsHostName == null)
                {
                    Logger.Error($"Ошибка заполнения пути для копирования с Revit-Server: ({pathFrom}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'");
                    return null;
                }
                try
                {
                    RevitServer server = new RevitServer(rsHostName, int.Parse(RevitVersion));
                    // Проверяю ссылку на конечный файл. Добавляю файл
                    if (pathFrom.ToLower().Contains("rvt"))
                    {
                        FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 4));
                        foreach (var model in folderContents.Models)
                        {
                            if (model.Name == pathParts[pathPartsLenght - 1])
                            {
                                fileFromPathes.Add($"RSN:{pathFrom}");
                                break;
                            }
                        }
                    }
                    // Значит ссылка на папку. Добавляю файлы
                    else
                    {
                        FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathPartsLenght - 3));
                        foreach (var model in folderContents.Models)
                        {
                            fileFromPathes.Add($"RSN:{pathFrom}{model.Name}");
                        }
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
