using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using RevitServerAPILib;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;


namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandNWExport : ExchangeService, IExternalCommand, IExecuteByUIApp
    {
        internal const string PluginName = "NWC: Экспорт";

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
                StartService(uiapp, revitDocExchangeEnum, PluginName);
            }
            catch (Exception ex)
            {
                PrintError(ex);

                return Result.Failed;
            }
            return Result.Succeeded;
        }

        private protected override string ExchangeFile(Application app, ModelPath modelPathFrom, DBConfigEntity configEntity, string rsn = "")
        {
            //Апкастинг в настройку для экспорта в NW
            if (configEntity is DBNWConfigData nwConfigData)
            {
                #region Анализ и открытие рабочих наборов
                IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(modelPathFrom);
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
                Document doc = null;
                try
                {
                    doc = app.OpenDocumentFile(modelPathFrom, _openOptions);
                }
                catch (Exception ex)
                {
                    Logger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). Нужно вмешаться человеку, " +
                        $"ошибка при открытии: {ex.Message}");
                    
                    return null;
                }

                #region Поиск и проверка вида для экспорта
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

                ElementId viewId = currentDoc3DViews.First().Id;

                var viewElemsColl = new FilteredElementCollector(doc, viewId)
                    .WhereElementIsNotElementType()
                    .Where(e => e.Category != null && e.Category.IsVisibleInUI)
                    .ToArray();
                    
                if (viewElemsColl.Length == 0)
                {
                    Logger.Error($"На вид с именем ({nwConfigData.ViewName}) НЕТ элементов для экспорта. Нужно вмешаться человеку");
                    return null;
                }

                exportOptions.ViewId = viewId;
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
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBNWConfigData)}");
        }
    }
}
