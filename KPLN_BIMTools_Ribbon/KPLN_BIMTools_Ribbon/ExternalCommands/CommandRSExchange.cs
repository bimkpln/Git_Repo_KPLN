using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    /// <summary>
    /// Плагин по импорту моделей с Revit-Server
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandRSExchange : ExchangeEnvironment, IExternalCommand, IExecuteByUIApp
    {
        public CommandRSExchange()
        {
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
            RevitEventWorker revitEventWorker = new RevitEventWorker(this, Logger, DBRevitDialogs);

            // Подписка на события
            RevitUIControlledApp.DialogBoxShowing += revitEventWorker.OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened += revitEventWorker.OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed += revitEventWorker.OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing += revitEventWorker.OnFailureProcessing;

            Logger.Info($"Старт обмена файлами с Revit-Server");

            // Подготовка к открытию
            SetOptions(WorksetConfigurationOption.CloseAllWorksets);

            // Копирую файлы по указанным путям
            foreach (DBRevitDocExchanges docExchanges in DBRevitDocExchanges)
            {
                _sourceProjectsName.Add(CurrentProjectDbService.GetDBProject_ByProjectId(docExchanges.ProjectId).Name);
                PrepareAndCopyFile(uiapp.Application, docExchanges.PathFrom, docExchanges.PathTo);
            }

            SendResultMsg("Плагин по обмену моделями с Revit-Server");
            Logger.Info($"Файлы успешно переданы на Revit-Server!");

            //Отписка от событий
            RevitUIControlledApp.DialogBoxShowing -= revitEventWorker.OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened -= revitEventWorker.OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed -= revitEventWorker.OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= revitEventWorker.OnFailureProcessing;

            revitEventWorker.Dispose();

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод подготовки к копированию и копированию
        /// </summary>
        private void PrepareAndCopyFile(Application app, string pathFrom, string pathTo)
        {
            List<string> fileFromPathes = PreparePathesToOpen(pathFrom);
            if (fileFromPathes.Count == 0)
            {
                Logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return;
            }

            // Проверяю, что это папка, если нет - то ревит-сервер
            if (Directory.Exists(pathTo))
                OpenAndCopyFile(app, fileFromPathes, pathFrom, pathTo);
            // Обрабатываю ревит-сервер
            else
                OpenAndCopyFile(app, fileFromPathes, pathFrom, pathTo, "RSN:");
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
                    CountSourceDocs++;
                    // Открываем документ по указанному пути
                    Document doc = app.OpenDocumentFile(
                        ModelPathUtils.ConvertUserVisiblePathToModelPath(currentPathFrom),
                        _openOptions);

                    if (doc != null)
                    {
                        string newPath = $"{rsn}{pathTo}\\{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}.rvt";
                        ModelPath newModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(newPath);

                        doc.SaveAs(newModelPath, _saveAsOptions);
                        CurrentDocName = doc.Title;
                        doc.Close();

                        CountProcessedDocs++;
                        Logger.Info($"Файл {newPath} успешно сохранен!\n");
                    }
                    else
                        Logger.Error($"Не удалось открыть Revit-документ ({currentPathFrom}). Нужно вмешаться человеку");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка обработки Revit-документа/файлов из папки ({pathFrom}):\n{ex.Message}");
            }
        }
    }
}
