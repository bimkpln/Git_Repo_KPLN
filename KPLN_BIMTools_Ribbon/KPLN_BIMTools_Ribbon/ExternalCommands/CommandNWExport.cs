using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExternalService;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandNWExport : ExchangeEnvironment, IExternalCommand, IExecuteByUIApp
    {
        public CommandNWExport()
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
            exportOptions = new NavisworksExportOptions();

            // Подготовка к открытию
            SetOptions(WorksetConfigurationOption.OpenAllWorksets);

            // Экспортирую файлы по настройкам из конфигов
            // Затычка
            DBRevitDocExchanges[] temp = new DBRevitDocExchanges[] 
            { 
                new DBRevitDocExchanges () 
                { 
                    ProjectId = 12, 
                    PathFrom = "\\\\stinproject.local\\project\\Жилые здания\\Речной порт Якутск\\10.Стадия_Р\\7.3.АУПТ\\1.RVT\\Архив\\ОВ_Модель.rvt",
                    PathTo = "\\\\stinproject.local\\project\\Жилые здания\\Речной порт Якутск\\10.Стадия_Р\\BIM\\1.Модели_для_проверки_Navisworks"
                }};
            foreach (DBRevitDocExchanges docExchanges in temp)
            //foreach (DBRevitDocExchanges docExchanges in DBRevitDocExchanges)
            {
                _sourceProjectsName.Add(CurrentProjectDbService.GetDBProject_ByProjectId(docExchanges.ProjectId).Name);
                PrepareAndExportFile(uiapp.Application, docExchanges.PathFrom, docExchanges.PathTo, exportOptions);
            }

            SendResultMsg("Плагин по экспорту в Navisworks");
            Logger.Info($"Файлы успешно экспортированы в Navisworks!");

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
        private void PrepareAndExportFile(Application app, string pathFrom, string pathTo, NavisworksExportOptions exportOptions)
        {
            List<string> fileFromPathes = PreparePathesToOpen(pathFrom);
            if (fileFromPathes.Count == 0)
            {
                Logger.Error($"Не удалось найти Revit-файлы из папки: {pathFrom}");
                return;
            }

            // Проверяю, что это папка, если нет - то ревит-сервер
            if (Directory.Exists(pathTo))
                ExportFile(app, fileFromPathes, pathFrom, pathTo, exportOptions);
            // Обрабатываю ревит-сервер
            else
                ExportFile(app, fileFromPathes, pathFrom, pathTo, exportOptions, "RSN:");
        }

        private void ExportFile(Application app, List<string> fileFromPathes, string pathFrom, string pathTo, NavisworksExportOptions exportOptions, string rsn = "")
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
                        string folderTo = $"{rsn}{pathTo}";
                        CurrentDocName = $"{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}.nwc";

                        doc.Export(folderTo, CurrentDocName, exportOptions);
                        doc.Close(false);

                        CountProcessedDocs++;
                        Logger.Info($"Файл {folderTo}\\{CurrentDocName} успешно экспортирован в Navisworks!\n");
                    }
                    else
                        Logger.Error($"Не удалось открыть Revit-документ ({currentPathFrom}). Нужно вмешаться человеку");
                }
            }
            catch (OptionalFunctionalityNotAvailableException nwEx)
            {
                Logger.Error($"На твоём компьютере отсутсвует плагин для экспорта в Navisworks. Тело ошибки: \n{nwEx.Message}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Ошибка обработки Revit-документа/файлов из папки ({pathFrom}):\n{ex.Message}");
            }
        }
    }
}
