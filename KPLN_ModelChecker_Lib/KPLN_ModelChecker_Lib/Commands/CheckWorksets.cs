using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Commands
{
    public sealed class CheckWorksets : AbstrCheck<CheckWorksets>
    {
        public CheckWorksets(UIApplication uiapp) : base(uiapp) { }

        public override Element[] GetElemsToCheck() => new FilteredElementCollector(CheckDoc).WhereElementIsNotElementType().ToArray();

        private protected override IEnumerable<CheckCommandError> CheckRElems(object[] objColl) => Enumerable.Empty<CheckCommandError>();

        private protected override IEnumerable<CheckerEntity> GetCheckerEntities(Element[] elemColl)
        {
            List<CheckerEntity> result = new List<CheckerEntity>();

            if (CheckDoc.IsWorkshared)
            {
                Workset[] worksets = new FilteredWorksetCollector(CheckDoc).OfKind(WorksetKind.UserWorkset).Where(w => w.IsOpen).ToArray();

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
                            result.Add(new CheckerEntity(
                                link,
                                "Ошибка рабочего набора",
                                "Связь находится в некорректном рабочем наборе",
                                "Для связей необходимо использовать именные рабочие наборы, которые начинаются с '00_' (если иное не указано в ВЕР для проекта)",
                                false));
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
                                result.Add(new CheckerEntity(
                                    impInstance,
                                    "Ошибка рабочего набора",
                                    "Связь находится в некорректном рабочем наборе",
                                    "Для DWG-связей необходимо использовать именной рабочий набор - '00_DWG' (если иное не указано в ВЕР для проекта)",
                                    false));
                            }
                        }
                        continue;
                    }

                    //Анализ уровней и осей
                    if ((element.Category.CategoryType == CategoryType.Annotation) & (element.GetType() == typeof(Autodesk.Revit.DB.Grid) | element.GetType() == typeof(Level)))
                    {
                        string wsName = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                        if (!wsName.ToLower().Contains("оси и уровни") & !wsName.ToLower().Contains("общие уровни и сетки"))
                        {
                            result.Add(new CheckerEntity(
                                element,
                                "Ошибка сеток",
                                $"Ось или уровень с ID: {element.Id} находится не в специальном рабочем наборе",
                                "Имя рабочего набора для осей и уровней - <..._Оси и уровни>",
                                false));
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
                                CheckerEntity entity = new CheckerEntity(
                                    element,
                                    "Ошибка мониторинговых элементов",
                                    $"Элементс с ID: {element.Id} находится не в специальном рабочем наборе",
                                    "Элементы с мониторингом (т.е. скопированные из других файлов) должны находится в рабочих наборах с приставкой '02'",
                                    true);

                                result.Add(entity);
                                continue;
                            }
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор связей 
                        else if (elemWSName.StartsWith("00")
                            | elemWSName.StartsWith("#")
                            && !elemWSName.ToLower().Contains("dwg"))
                        {
                            CheckerEntity entity = new CheckerEntity(
                                element,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для связей",
                                string.Empty,
                                true);

                            result.Add(entity);
                            continue;
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор для сеток
                        else if (elemWSName.ToLower().Contains("оси и уровни")
                            | elemWSName.ToLower().Contains("общие уровни и сетки"))
                        {
                            CheckerEntity entity = new CheckerEntity(
                                element,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для осей и уровней",
                                string.Empty,
                                true);

                            result.Add(entity);
                            continue;
                        }

                        // Проверка остальных моделируемых элементов на рабочий набор для связей
                        else if (elemWSName.StartsWith("02"))
                        {
                            CheckerEntity entity = new CheckerEntity(
                                element,
                                "Ошибка элементов",
                                $"Элементс с ID: {element.Id} находится в рабочем наборе для элементов с монитирнгом",
                                string.Empty,
                                true);

                            result.Add(entity);
                            continue;
                        }
                    }
                }
            }

            return result;
        }
    }
}
