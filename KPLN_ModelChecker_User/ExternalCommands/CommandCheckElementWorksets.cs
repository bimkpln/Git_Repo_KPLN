using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    internal class CommandCheckElementWorksets : AbstrCheckCommand<CommandCheckElementWorksets>, IExternalCommand
    {
        public CommandCheckElementWorksets() : base()
        {
        }

        internal CommandCheckElementWorksets(ExtensibleStorageEntity esEntity) : base(esEntity)
        {
        }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return ExecuteByUIApp(commandData.Application);
        }

        public override Result ExecuteByUIApp(UIApplication uiapp)
        {
            _uiApp = uiapp;

            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Получаю коллекцию элементов для анализа
            Element[] allModelElems = new FilteredElementCollector(doc).WhereElementIsNotElementType().ToArray();

            #region Проверяю и обрабатываю элементы
            WPFEntity[] wpfColl = CheckCommandRunner(doc, allModelElems);
            OutputMainForm form = ReportCreatorAndDemonstrator(doc, wpfColl);
            if (form != null) form.Show();
            else return Result.Cancelled;
            #endregion

            return Result.Succeeded;
        }

        private protected override IEnumerable<CheckCommandError> CheckElements(Document doc, object[] elemColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl)
        {
            List<WPFEntity> result = new List<WPFEntity>();

            if (doc.IsWorkshared)
            {
                List<Workset> worksets = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToList();
                if (worksets.Where(w => !w.IsOpen).Any())
                    throw new CheckerException("Необходимо открыть все рабочие наборы!");

                foreach (Element element in elemColl)
                {
                    // Игнор безкатегорийных эл-в
                    if (element.Category == null) { continue; }

                    // Анализ связей
                    if (element is RevitLinkInstance link)
                    {
                        string[] separators = { ".rvt : " };
                        string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                        if (nameSubs.Length > 3) continue;

                        string wsName = link.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                        if (!wsName.StartsWith("00") && !wsName.StartsWith("#"))
                        {
                            result.Add(new WPFEntity(
                                link,
                                CheckStatus.Error,
                                "Ошибка рабочего набора",
                                "Связь находится в некорректном рабочем наборе",
                                false,
                                false,
                                "Для связей необходимо использовать именные рабочие наборы, которые начинаются с '00_'"));
                        }
                        continue; 
                    }
                    else if (element is ImportInstance impInstance)
                    {
                        // DWG может по разному импортировать связью. Те, что прикрепляются к уровню - могут иметь разный рабочий набор
                        if (impInstance.IsLinked && !impInstance.ViewSpecific)
                        {
                            string wsName = impInstance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                            if (!wsName.StartsWith("00") && !wsName.StartsWith("#"))
                            {
                                result.Add(new WPFEntity(
                                    impInstance,
                                    CheckStatus.Error,
                                    "Ошибка рабочего набора",
                                    "Связь находится в некорректном рабочем наборе",
                                    false,
                                    false,
                                    "Для связей необходимо использовать именные рабочие наборы, которые начинаются с '00_'"));
                            }
                        }
                        continue;
                    }

                    //Анализ уровней и осей
                    if ((element.Category.CategoryType == CategoryType.Annotation) & (element.GetType() == typeof(Grid) | element.GetType() == typeof(Level)))
                    {
                        string wsName = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                        if (!wsName.ToLower().Contains("оси и уровни") & !wsName.ToLower().Contains("общие уровни и сетки"))
                        {
                            result.Add(new WPFEntity(
                                element,
                                CheckStatus.Error,
                                "Ошибка сеток",
                                $"Ось или уровень с ID: {element.Id} находится не в специальном рабочем наборе",
                                false,
                                false,
                                "Имя рабочего набора для осей и уровней - <..._Оси и уровни>"
                                ));
                        }
                    }

                    //Анализ моделируемых элементов
                    if (element.Category.CategoryType == CategoryType.Model
                        // Есть внутренняя ошибка Revit, когда появляются компоненты легенды, которые нигде не размещены, и у них редактируемый рабочий набор. Вручную такой элемент - создать НЕВОЗМОЖНО
                        && (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_PreviewLegendComponents
                        // Игнор зон ОВК
                        && (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_HVAC_Zones
                        // Игнор набора характеристик материалов
                        && (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_PropertySet
                        // Игнор эскизов
                        && (BuiltInCategory)element.Category.Id.IntegerValue != BuiltInCategory.OST_SketchLines)
                    {
                        string elemWSName = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                        Workset elemWS = worksets.Where(w => w.Name.Equals(elemWSName)).FirstOrDefault();
                        if (elemWS == null) { continue; }

                        // Проверка замонитренных моделируемых элементов
                        if (element.GetMonitoredLinkElementIds().Count() > 0)
                        {
                            if (!elemWSName.StartsWith("02"))
                            {
                                WPFEntity entity = new WPFEntity(
                                    element,
                                    CheckStatus.Error,
                                    "Ошибка мониторинговых элементов",
                                    $"Элементс с ID: {element.Id} находится не в специальном рабочем наборе",
                                    true,
                                    false,
                                    "Элементы с мониторингом (т.е. скопированные из других файлов) должны находится в рабочих наборах с приставкой '02'"
                                    );
                                entity.PrepareZoomGeometryExtension(element.get_BoundingBox(null));
                                result.Add(entity);
                                continue;
                            }
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор связей 
                        else if (elemWSName.StartsWith("00")
                            | elemWSName.StartsWith("#")
                            && !elemWSName.ToLower().Contains("dwg"))
                        {
                            WPFEntity entity = new WPFEntity(
                                element,
                                CheckStatus.Error,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для связей",
                                true,
                                false
                                );
                            entity.PrepareZoomGeometryExtension(element.get_BoundingBox(null));
                            result.Add(entity);
                            continue;
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор для сеток
                        else if (elemWSName.ToLower().Contains("оси и уровни")
                            | elemWSName.ToLower().Contains("общие уровни и сетки"))
                        {
                            WPFEntity entity = new WPFEntity(
                                element,
                                CheckStatus.Error,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для осей и уровней",
                                true,
                                false
                                );
                            entity.PrepareZoomGeometryExtension(element.get_BoundingBox(null));
                            result.Add(entity);
                            continue;
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор для связей
                        else if (elemWSName.StartsWith("02"))
                        {
                            WPFEntity entity = new WPFEntity(
                                element,
                                CheckStatus.Error,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для элементов с монитирнгом",
                                true,
                                false
                                );
                            entity.PrepareZoomGeometryExtension(element.get_BoundingBox(null));
                            result.Add(entity);
                            continue;
                        }
                    }
                }
            }
            else
                throw new CheckerException("Только для файлов, в которых есть рабочие наборы");

            return result;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}
