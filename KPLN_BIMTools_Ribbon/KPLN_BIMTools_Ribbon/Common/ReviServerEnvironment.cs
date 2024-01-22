using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Общий класс по работе с Revit-Server
    /// </summary>
    public class ReviServerEnvironment
    {
        private List<DBRevitDialog> _dbRevitDialogs;

        private DBUser _dBUser;

        public ReviServerEnvironment(UIControlledApplication application, Logger logger)
        {
            RevitUIControlledApp = application;
            Logger = logger;
        }

        public ReviServerEnvironment()
        {
        }

        internal protected static UIControlledApplication RevitUIControlledApp { get; private set; }

        internal protected static Logger Logger { get; private set; }

        /// <summary>
        /// Счтетчик успешно отработанных процессов
        /// </summary>
        internal protected int CountProcessedDocs { get; protected set; } = 0;

        /// <summary>
        /// Список диалогов из БД
        /// </summary>
        internal protected List<DBRevitDialog> DBRevitDialogs
        {
            get
            {
                if (_dbRevitDialogs == null)
                {
                    RevitDialogDbService dialogDbService = (RevitDialogDbService)new CreatorRevitDialogtDbService().CreateService();
                    _dbRevitDialogs = dialogDbService.GetDBRevitDialogs().ToList();
                }
                return _dbRevitDialogs;
            }
        }

        /// <summary>
        /// Ссылка на текущего пользователя из БД
        /// </summary>
        internal DBUser CurrentDBUser
        {
            get
            {
                if (_dBUser == null)
                {
                    UserDbService userDbService = (UserDbService)new CreatorUserDbService().CreateService();
                    _dBUser = userDbService.GetCurrentDBUser();
                }

                return _dBUser;
            }
        }

        /// <summary>
        /// Метод открытия и копирования файла по новому пути
        /// </summary>
        internal bool OpenAndCopyFile(Application app, OpenOptions openOptions, SaveAsOptions saveAsOptions, string filePath, string newFilePath)
        {
            // Проверяем наличие папки
            if (!Directory.Exists(newFilePath))
            {
                // Создаем папку, если она отсутствует
                try
                {
                    Directory.CreateDirectory(newFilePath);
                }
                catch (Exception ex)
                {
                    Logger.Error($"Ошибка при создании папки ({newFilePath}):\n{ex.Message}");
                }
            }

            try
            {
                // Открываем документ по указанному пути
                Document doc = app.OpenDocumentFile(
                    ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath), 
                    openOptions);

                if (doc != null)
                {
                    doc.SaveAs($"{newFilePath}\\{doc.Title}.rvt", saveAsOptions);
                    doc.Close();
                    Logger.Info($"Файл успешно сохранен!");
                }
                else
                {
                    Logger.Error($"Не удалось открыть Revit-документ ({filePath}). Нужно вмешаться человеку");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка открытия Revit-документа ({filePath}):\n{ex.Message}");
                return false;
            }

            return true;
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
