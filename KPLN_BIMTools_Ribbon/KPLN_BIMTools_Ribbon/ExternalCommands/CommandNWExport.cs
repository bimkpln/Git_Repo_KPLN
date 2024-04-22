using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandNWExport : ExchangeService, IExternalCommand, IExecuteByUIApp
    {
        public CommandNWExport()
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application, RevitDocExchangeEnum.Navisworks);
        }

        public Result ExecuteByUIApp(UIApplication uiapp, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            try
            {
                if (!StartService(uiapp, revitDocExchangeEnum))
                    return Result.Cancelled;
            }
            catch (Exception ex)
            {
                PrintError(ex);

                return Result.Failed;
            }
            return Result.Succeeded;
        }

        private protected override string ExchangeFile(Application app, string fileFromPath, DBConfigEntity configEntity, string rsn = "")
        {
            //Апкастинг в настройку для экспорта в NW
            if (configEntity is DBNWConfigData nwConfigData)
            {
                ModelPath docModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(fileFromPath);

                #region Анализ и открытие рабочих наборов
                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(docModelPath);
                IList<WorksetId> worksetIds = new List<WorksetId>();

                string[] wsExceptions = nwConfigData.WorksetToCloseNamesStartWith.Split('~');
                foreach (WorksetPreview worksetPrev in worksets)
                {
                    if (nwConfigData.WorksetToCloseNamesStartWith.Count() == 0)
                        worksetIds.Add(worksetPrev.Id);
                    else if (!wsExceptions.Any(name => worksetPrev.Name.StartsWith(name)))
                        worksetIds.Add(worksetPrev.Id);
                }
                SetOpenOptions(worksetIds);
                #endregion

                #region Устанавливаем параметры экспорта в Navisworks
                NavisworksExportOptions exportOptions = new NavisworksExportOptions
                {
                    FacetingFactor = nwConfigData.FacetingFactor,
                    ConvertElementProperties = nwConfigData.ConvertElementProperties,
                    ExportLinks = nwConfigData.ExportLinks,
                    FindMissingMaterials = nwConfigData.FindMissingMaterials,
                    ExportScope = nwConfigData.ExportScope,
                    DivideFileIntoLevels = nwConfigData.DivideFileIntoLevels,
                    ExportRoomGeometry = nwConfigData.ExportRoomGeometry,
                };
                #endregion

                // Открываем документ по указанному пути
                Document doc = app.OpenDocumentFile(docModelPath, _openOptions);

                if (doc != null)
                {
                    #region Поиск вида для экспорта
                    IEnumerable<View3D> currentDoc3DViews = new FilteredElementCollector(doc)
                        .OfClass(typeof(View3D))
                        .WhereElementIsNotElementType()
                        .Where(e => e.Name.Equals(nwConfigData.ViewName))
                        .Select(e => e as View3D);
                    if (currentDoc3DViews.Count() == 0)
                    {
                        Logger.Error($"Не удалось найти вид с именем ({nwConfigData.ViewName}). Либо конфигурация не верная, либо такого вида в проекте нет. Нужно вмешаться человеку");
                        return null;
                    }
                    exportOptions.ViewId = currentDoc3DViews.First().Id;
                    #endregion

                    #region Экспорт в Navisworks
                    string folderTo = $"{rsn}{nwConfigData.PathTo}";
                    CurrentDocName = $"{doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0]}{nwConfigData.NavisDocPostfix}.nwc";

                    doc.Export(folderTo, CurrentDocName, exportOptions);
                    doc.Close(false);

                    return $"{folderTo}\\{CurrentDocName}";
                    #endregion
                }
                else
                    Logger.Error($"Не удалось открыть Revit-документ ({fileFromPath}). Нужно вмешаться человеку");
            }
            else
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBNWConfigData)}");

            return null;
        }
    }
}
