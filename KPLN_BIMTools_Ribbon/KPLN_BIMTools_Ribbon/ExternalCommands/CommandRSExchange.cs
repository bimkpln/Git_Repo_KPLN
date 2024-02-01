using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    /// <summary>
    /// Плагин по импорту моделей с Revit-Server
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandRSExchange : IExternalCommand, IExcecuteByUIApp
    {
        private static RevitDocExchangestDbService _revitDocExchangestDbService;
        private static UserDbService _userDbService;
        private static RevitDialogDbService _dialogDbService;
        private static ProjectDbService _projectDbService;
        private static UIControlledApplication _revitUIControlledApp;
        private static NLog.Logger _logger;
        private static string _revitVersion;

        private DBUser _dBUser;
        private OpenOptions _openOptions;
        private SaveAsOptions _saveAsOptions;
        /// <summary>
        /// Счтетчик файлов для копирования
        /// </summary>
        private int _countSourceDocs = 0;
        /// <summary>
        /// Счтетчик успешно отработанных процессов
        /// </summary>
        private int _countProcessedDocs = 0;
        private HashSet<string> _sourceProjectsName = new HashSet<string>();
        private string _currentDocName;

        public CommandRSExchange()
        {
        }

        internal static RevitDocExchangestDbService CurrentRevitDocExchangestDbService
        {
            get
            {
                if (_revitDocExchangestDbService == null)
                    _revitDocExchangestDbService = (RevitDocExchangestDbService)new CreatorRevitDocExchangesDbService().CreateService();
                return _revitDocExchangestDbService;
            }
        }

        internal static UserDbService CurrentUserDbService
        {
            get
            {
                if (_userDbService == null)
                    _userDbService = (UserDbService)new CreatorUserDbService().CreateService();
                return _userDbService;
            }
        }

        internal static RevitDialogDbService CurrentRevitDialogDbService
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

        /// <summary>
        /// Список диалогов из БД
        /// </summary>
        internal DBRevitDialog[] DBRevitDialogs
        {
            get => CurrentRevitDialogDbService.GetDBRevitDialogs().ToArray();
        }

        /// <summary>
        /// Ссылка на массив документов для обмена данными
        /// </summary>
        internal DBRevitDocExchanges[] DBRevitDocExchanges
        {
            get => CurrentRevitDocExchangestDbService.GetDBRevitActiveDocExchanges().ToArray();
        }


        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                    _dBUser = CurrentUserDbService.GetCurrentDBUser();

                return _dBUser;
            }
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public Result ExecuteByUIApp(UIApplication uiapp)
        {
            // Подписка на события
            _revitUIControlledApp.DialogBoxShowing += OnDialogBoxShowing;
            _revitUIControlledApp.ControlledApplication.DocumentOpened += OnDocumentOpened;
            _revitUIControlledApp.ControlledApplication.DocumentClosed += OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing += OnFailureProcessing;

            _logger.Info($"Старт обмена файлами с Revit-Server");

            // Подготовка к открытию
            SetOptions();

            // Копирую файлы по указанным путям
            foreach (DBRevitDocExchanges docExchanges in DBRevitDocExchanges)
            {
                _sourceProjectsName.Add(CurrentProjectDbService.GetDBProject_ByProjectId(docExchanges.ProjectId).Name);
                PrepareAndCopyFile(uiapp.Application, docExchanges.PathFrom, docExchanges.PathTo);
            }

            SendResultMsg("Плагин по обмену моделями с Revit-Server");
            _logger.Info($"Файлы успешно переданы на Revit-Server!");

            //Отписка от событий
            _revitUIControlledApp.DialogBoxShowing -= OnDialogBoxShowing;
            _revitUIControlledApp.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            _revitUIControlledApp.ControlledApplication.DocumentClosed -= OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= OnFailureProcessing;

            return Result.Succeeded;
        }

        internal static void SetStaticEnvironment(UIControlledApplication application, NLog.Logger logger, string revitVersion)
        {
            _revitUIControlledApp = application;
            _logger = logger;
            _revitVersion = revitVersion;
        }

        /// <summary>
        /// Метод подготовки к копированию и копированию
        /// </summary>
        private void PrepareAndCopyFile(Application app, string pathFrom, string pathTo)
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
                    _logger.Error($"Ошибка заполнения пути для копирования с Revit-Server: ({pathFrom}). Путь должен быть в формате '\\\\HOSTNAME\\PATH'");
                    return;
                }
                try
                {
                    RevitServer server = new RevitServer(rsHostName, int.Parse(_revitVersion));
                    FolderContents folderContents = server.GetFolderContents(string.Join("\\", pathParts, 3, pathParts.Length - 3));
                    foreach (var model in folderContents.Models)
                    {
                        fileFromPathes.Add($"RSN:{pathFrom}{model.Name}");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"Ошибка открытия Revit-Server ({pathFrom}):\n{ex.Message}");
                    return;
                }

            }

            if (fileFromPathes.Count == 0)
            {
                _logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return;
            }

            // Проверяю, что это папка, если нет - то ревит-сервер
            if (Directory.Exists(pathTo))
            {
                OpenAndCopyFile(app, fileFromPathes, pathFrom, pathTo);
            }
            // Обрабатываю ревит-сервер
            else
            {
                OpenAndCopyFile(app, fileFromPathes, pathFrom, pathTo, "RSN:");
            }
        }

        /// <summary>
        /// Метод открытия и копирования файла по новому пути
        /// </summary>
        private void OpenAndCopyFile(Application app, List<string> fileFromPathes, string pathFrom, string pathTo, string rsn = "")
        {
            try
            {
                foreach (string currentPathFrom in fileFromPathes)
                {
                    _countSourceDocs++;
                    // Открываем документ по указанному пути
                    Document doc = app.OpenDocumentFile(
                        ModelPathUtils.ConvertUserVisiblePathToModelPath(currentPathFrom),
                        _openOptions);

                    if (doc != null)
                    {
                        string newPath = $"{rsn}{pathTo}\\{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}.rvt";
                        ModelPath newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath);

                        doc.SaveAs(newModelPath, _saveAsOptions);
                        _currentDocName = doc.Title;
                        doc.Close();

                        _countProcessedDocs++;
                        _logger.Info($"Файл {newPath} успешно сохранен!\n");
                    }
                    else
                    {
                        _logger.Error($"Не удалось открыть Revit-документ ({currentPathFrom}). Нужно вмешаться человеку");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Ошибка обработки Revit-документа/файлов из папки ({pathFrom}):\n{ex.Message}");
            }
        }

        /// <summary>
        /// Подготовка опций к открытию и сохранению
        /// </summary>
        private void SetOptions()
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
        private void SendResultMsg(string moduleName)
        {
            if (_countProcessedDocs < _countSourceDocs)
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано с ошибками.\n" +
                        $"Метрик производительности: Выгружено {_countProcessedDocs} из {_countSourceDocs} файлов, для проекта/-ов: {string.Join(", ", _sourceProjectsName)}\n" +
                        $"Ошибки: См. файл логов у пользователя {CurrentDBUser.Surname} {CurrentDBUser.Name}.\n" +
                        $"Путь к логам у пользователя: C:\\TEMP\\KPLN_Logs\\2023");
            }
            else
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано без ошибок.\n" +
                        $"Метрик производительности: Обработано {_countProcessedDocs} из {_countSourceDocs} файлов, для проекта/-ов: {string.Join(", ", _sourceProjectsName)}");
            }
        }
        /// <summary>
        /// Событие на открытие документа
        /// </summary>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            _logger.Info($"Начало работы с файлом: {args.Document.PathName}");
        }

        /// <summary>
        /// Событие на закрытие документа
        /// </summary>
        private void OnDocumentClosed(object sender, DocumentClosedEventArgs args)
        {
            _logger.Info($"Конец работы с файлом: {_currentDocName}");
        }

        /// <summary>
        /// Обработка события всплывающего окна Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            _logger.Info($"Появилось окно {args.DialogId}");

            if (args.Cancellable)
            {
                args.Cancel();
                _logger.Info($"Окно {args.DialogId} успешно закрыто, благодаря стандартной возможности закрывания (Cancellable) данного окна");
            }
            else
            {
                DBRevitDialog currentDBDialog = DBRevitDialogs.FirstOrDefault(rd => args.DialogId.Contains(rd.DialogId));
                if (currentDBDialog != null)
                {
                    if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                    {
                        if (args.OverrideResult((int)taskDialogResult))
                            _logger.Info($"Окно {args.DialogId} успешно закрыто. Была применена команда {currentDBDialog.OverrideResult}");
                        else
                            _logger.Error($"Окно {args.DialogId} не удалось обработать. Была применена команда {currentDBDialog.OverrideResult}, но она не сработала!");
                    }

                    _logger.Error($"Не удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!");

                }
                else
                {
                    _logger.Error($"Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека");
                }
            }
        }

        private void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            // Пока не понятно, нужна ли реализация
        }

    }
}
