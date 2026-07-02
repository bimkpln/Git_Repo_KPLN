using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_DBWorker.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
            if (!(configEntity is DBNWConfigData nwConfigData))
                throw new Exception($"Скинь разработчику: Не удалось совершить корректный апкастинг из {nameof(DBConfigEntity)} в {nameof(DBNWConfigData)}");


            // Проверка на Revit Server
            if (!string.IsNullOrWhiteSpace(rsn))
            {
                Module.CurrentLogger.Error("Экспорт NMWC на Revit Server не поддерживается. Нужно указать обычную сетевую/локальную папку.");
                return null;
            }


            #region Анализ и открытие рабочих наборов
            IList<WorksetPreview> worksets = WorksharingUtils.GetUserWorksetInfo(modelPathFrom);
            IList<WorksetId> worksetIds = new List<WorksetId>();

            StringBuilder openedWSSB = new StringBuilder();
            foreach (WorksetPreview worksetPrev in worksets)
            {
                if (!WSName_IsMatchByRules(worksetPrev.Name, nwConfigData.WorksetToCloseNamesStartWith))
                {
                    worksetIds.Add(worksetPrev.Id);
                    openedWSSB.Append($"{worksetPrev.Name}, ");
                }
            }
            SetOpenOptions(worksetIds);

            // Логирую список закрытых РН
            Module.CurrentLogger.Info($"Список открываемых РН в файле {ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}: {openedWSSB.ToString().TrimEnd(new char[] { ',', ' ' })}");
            #endregion

            
            // Открываем документ по указанному пути
            Document doc = null;
            try
            {
                try
                {
                    doc = app.OpenDocumentFile(modelPathFrom, _openOptions);
                }
                catch (Exception ex)
                {
                    Module.CurrentLogger.Error($"Не удалось открыть Revit-документ ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). Нужно вмешаться человеку, " +
                        $"ошибка при открытии: {ex.Message}");

                    return null;
                }

                #region Поиск и проверка вида для экспорта
                IEnumerable<View3D> currentDoc3DViews = new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D))
                    .WhereElementIsNotElementType()
                    .Where(e => e.Name.Equals(nwConfigData.ViewName))
                    .Select(e => e as View3D);
                if (!currentDoc3DViews.Any())
                {
                    Module.CurrentLogger.Error($"Не удалось найти вид с именем ({nwConfigData.ViewName}). Либо конфигурация не верная, либо такого вида в проекте нет. Нужно вмешаться человеку");
                    return null;
                }

                ElementId viewId = currentDoc3DViews.First().Id;

                var viewElemsColl = new FilteredElementCollector(doc, viewId)
                    .WhereElementIsNotElementType()
                    .Where(e =>
                        e.Category != null
                        && e.Category.CategoryType == CategoryType.Model
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                        && (e.Category.Id.IntegerValue == (int)BuiltInCategory.OST_RvtLinks || e.Category.IsVisibleInUI))
#else
                    && (e.Category.Id.Value == (long)BuiltInCategory.OST_RvtLinks || e.Category.IsVisibleInUI))
#endif
                    .ToArray();

                if (viewElemsColl.Length == 0)
                {
                    Module.CurrentLogger.Error($"На виде с именем ({nwConfigData.ViewName}) НЕТ элементов для экспорта. Нужно вмешаться человеку");
                    return null;
                }
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
                    ViewId = viewId,
                };
                #endregion


                #region Экспорт в Navisworks
                string folderTo = nwConfigData.PathTo;
                if (!Directory.Exists(folderTo))
                    Directory.CreateDirectory(folderTo);

                string docTitle = doc.Title.Split(new[] { "_отсоединено" }, StringSplitOptions.None)[0];
                CurrentDocName = $"{docTitle}{nwConfigData.NavisDocPostfix}.nwc";

                doc.Export(folderTo, CurrentDocName, exportOptions);

                return $"{folderTo}\\{CurrentDocName}";
                #endregion
            }
            catch (Exception ex)
            {
                Module.CurrentLogger.Error(
                    $"Не удалось экспортировать NWC ({ModelPathUtils.ConvertModelPathToUserVisiblePath(modelPathFrom)}). " +
                    $"Ошибка: {ex.Message}");

                return null;
            }
            finally
            {
                if (doc != null && doc.IsValidObject)
                {
                    try
                    {
                        doc.Close(false);
                    }
                    catch (Exception ex)
                    {
                        Module.CurrentLogger.Error($"Не удалось закрыть документ после NWC-экспорта. Ошибка: {ex.Message}");
                    }
                }
            }
        }
    }
}
