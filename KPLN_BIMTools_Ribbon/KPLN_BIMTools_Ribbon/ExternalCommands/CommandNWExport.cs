using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using RevitServerAPILib;
using System.Collections.Generic;
using System.IO;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    internal class CommandNWExport : ExchangeEnvironment, IExternalCommand, IExecuteByUIApp
    {
        public CommandNWExport()
        {
        }

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

            Logger.Info($"Старт экспорта в Navisworks");

            // Тут нужно настроить парсинг из файла настроек
            NWSettings settings = new NWSettings()
            {
                FacetingFactor = 1.0,
                ConvertElementProperties = true,
                ExportLinks = true,
                FindMissingMaterials = true,
                ExportScope = NavisworksExportScope.View,
                DivideFileIntoLevels = true,
                ExportRoomGeometry = true,
            };

            // Устанавливаем параметры экспорта
            NavisworksExportOptions exportOptions = new NavisworksExportOptions
            {
                FacetingFactor = settings.FacetingFactor,
                ConvertElementProperties = settings.ConvertElementProperties,
                ExportLinks = settings.ExportLinks,
                FindMissingMaterials = settings.FindMissingMaterials,
                ExportScope = settings.ExportScope,
                DivideFileIntoLevels = settings.DivideFileIntoLevels,
                ExportRoomGeometry = settings.ExportRoomGeometry,
            };

            // Копирую файлы по указанным путям
            foreach (DBRevitDocExchanges docExchanges in DBRevitDocExchanges)
            {
                _sourceProjectsName.Add(CurrentProjectDbService.GetDBProject_ByProjectId(docExchanges.ProjectId).Name);
                PrepareAndExportFile(uiapp.Application, docExchanges.PathFrom, docExchanges.PathTo);
            }

            SendResultMsg("Плагин по экспорту в Navisworks");
            Logger.Info($"Файлы успешно экспортированы!");

            //Отписка от событий
            RevitUIControlledApp.DialogBoxShowing -= revitEventWorker.OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened -= revitEventWorker.OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed -= revitEventWorker.OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= revitEventWorker.OnFailureProcessing;

            revitEventWorker.Dispose();

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод подготовки и экспорта файла
        /// </summary>
        private void PrepareAndExportFile(Application app, string pathFrom, string pathTo)
        {
            List<string> fileFromPathes = PreparePathesToOpen(pathFrom);
            if (fileFromPathes.Count == 0)
            {
                Logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return;
            }

            // Проверяю, что это папка, если нет - то ревит-сервер
            if (Directory.Exists(pathTo))
            {
                ExportFile(app, fileFromPathes, pathFrom, pathTo);
            }
            // Обрабатываю ревит-сервер
            else
            {
                ExportFile(app, fileFromPathes, pathFrom, pathTo, "RSN:");
            }
        }

        private void ExportFile(Application app, List<string> fileFromPathes, string pathFrom, string pathTo, string rsn = "")
        {

        }
    }
}
