using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using NLog;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Общий класс по работе с Revit-Server
    /// </summary>
    public class RevitServerEnvironment1
    {
        private static RevitDocExchangestDbService _revitDocExchangestDbService;
        private static UserDbService _userDbService;
        private static RevitDialogDbService _dialogDbService;
        private DBUser _dBUser;

        public RevitServerEnvironment1()
        {
        }

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

        internal protected static UIControlledApplication RevitUIControlledApp { get; private set; }

        internal protected static NLog.Logger Logger { get; private set; }
        
        internal protected static string RevitVersion { get; private set; }

        /// <summary>
        /// Счтетчик успешно отработанных процессов
        /// </summary>
        internal protected int CountProcessedDocs { get; protected set; } = 0;

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

        internal protected OpenOptions OpenOptions { get; private set; }
        
        internal protected SaveAsOptions SaveAsOptions { get; private set; }

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

        internal static void SetStaticEnvironment(UIControlledApplication application, NLog.Logger logger, string revitVersion)
        {
            RevitUIControlledApp = application;
            Logger = logger;
            RevitVersion = revitVersion;
        }

        /// <summary>
        /// Метод открытия и копирования файла по новому пути
        /// </summary>
        internal bool OpenAndCopyFile(Application app, string pathFrom, string pathTo)
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
                    return false;
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
                    return false;
                }

            }

            if (fileFromPathes.Count == 0)
            {
                Logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return false;
            }
            
            // Проверяем наличие папки для копирования
            if (!Directory.Exists(pathTo))
            {
                // Создаем папку, если она отсутствует
                try
                {
                    Directory.CreateDirectory(pathTo);
                }
                catch (Exception ex)
                {
                    Logger.Error($"Ошибка при создании папки ({pathTo}):\n{ex.Message}");
                }
            }

            try
            {
                foreach (string currentPathFrom in fileFromPathes)
                {
                    // Открываем документ по указанному пути
                    Document doc = app.OpenDocumentFile(
                        ModelPathUtils.ConvertUserVisiblePathToModelPath(currentPathFrom),
                        OpenOptions);

                    if (doc != null)
                    {
                        string newPath = $"{pathTo}\\{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}.rvt";
                        doc.SaveAs(newPath, SaveAsOptions);
                        doc.Close();
                        Logger.Info($"Файл {newPath} успешно сохранен!");
                    }
                    else
                    {
                        Logger.Error($"Не удалось открыть Revit-документ ({currentPathFrom}). Нужно вмешаться человеку");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка открытия Revit-документа/файлов из папки ({pathFrom}):\n{ex.Message}");
                return false;
            }

            return true;
        }

        /// <summary>
        /// Подготовка опций к открытию и сохранению
        /// </summary>
        internal void SetOptions()
        {
            OpenOptions = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets };
            OpenOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

            SaveAsOptions = new SaveAsOptions() { OverwriteExistingFile = true };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions() { SaveAsCentral = true };
            SaveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);
        }

        /// <summary>
        /// Отправка результата пользователю в месенджер
        /// </summary>
        internal void SendResultMsg(string moduleName, int filesCount)
        {
            if (CountProcessedDocs < filesCount)
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано с ошибками.\n" +
                        $"Метрик производительности: Выгружено {CountProcessedDocs} из {filesCount} файлов.\n" +
                        $"Ошибки: См. файл логов у пользователя {CurrentDBUser.Surname} {CurrentDBUser.Name}.\n" +
                        $"Путь к логам у пользователя: C:\\TEMP\\KPLN_Logs\\2023");
            }
            else
            {
                BitrixMessageSender.SendErrorMsg_ToBIMChat(
                        $"Модуль: {moduleName}\n" +
                        $"Статус: Отработано без ошибок.\n" +
                        $"Метрик производительности: Обработано {CountProcessedDocs} из {filesCount} файлов.");
            }
        }
        /// <summary>
        /// Событие на открытие документа
        /// </summary>
        internal void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Logger.Info($"Начало работы с файлом: {args.Document.PathName}");
        }

        /// <summary>
        /// Событие на закрытие документа
        /// </summary>
        internal void OnDocumentClosed(object sender, DocumentClosedEventArgs args)
        {
            if (sender is Document doc)
                Logger.Info($"Конец работы с файлом: {doc.PathName}\n");

            Logger.Error($"Ошибка приведения типа sender'а к Document\n");
        }

        /// <summary>
        /// Обработка события всплывающего окна Ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            Logger.Info($"Появилось окно {args.DialogId}");

            if (args.Cancellable)
            {
                args.Cancel();
                Logger.Info($"Окно {args.DialogId} успешно закрыто, благодаря стандартной возможности закрывания (Cancellable) данного окна");
            }
            else
            {
                DBRevitDialog currentDBDialog = DBRevitDialogs.FirstOrDefault(rd => args.DialogId.Contains(rd.DialogId));
                if (currentDBDialog != null)
                {
                    if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                    {
                        if (args.OverrideResult((int)taskDialogResult))
                            Logger.Info($"Окно {args.DialogId} успешно закрыто. Была применена команда {currentDBDialog.OverrideResult}");
                        else
                            Logger.Error($"Окно {args.DialogId} не удалось обработать. Была применена команда {currentDBDialog.OverrideResult}, но она не сработала!");
                    }

                    Logger.Error($"Не удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!");

                }
                else
                {
                    Logger.Error($"Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека");
                }
            }
        }

        internal void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            // Пока не понятно, нужна ли реализация
        }
    }
}
