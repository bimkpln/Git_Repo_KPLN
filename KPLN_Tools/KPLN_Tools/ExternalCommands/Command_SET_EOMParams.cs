using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    public class Command_SET_EOMParams : IExternalCommand
    {
        private const double FEET_TO_METERS = 0.3048;
        private const int ROUNDING_PRECISION = 2;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            var stopwatch = Stopwatch.StartNew();
            var failedElements = new List<Tuple<string, ElementId>>();
            var elementsToProcess = CollectElements(doc);

            using (var tx = new Transaction(doc, "Заполнение параметров лотков и воздуховодов"))
            {
                tx.Start();
                ProcessAllElements(elementsToProcess, failedElements);
                tx.Commit();
            }

            stopwatch.Stop();
            return DisplayResults(failedElements, uidoc, stopwatch.Elapsed);
        }
        private static List<Element> CollectElements(Document doc)
        {
            var categoryFilters = new List<ElementFilter>
            {
                new ElementCategoryFilter(BuiltInCategory.OST_CableTray),
                new ElementCategoryFilter(BuiltInCategory.OST_CableTrayFitting),
                new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)
            };

            return new FilteredElementCollector(doc)
                .WherePasses(new LogicalOrFilter(categoryFilters))
                .Where(e => e.Category != null && IsValidElement(e))
                .ToList();
        }
        private enum ElementType
        {
            None,
            CableTray,
            CableTrayFitting,
            DuctOGK,
            DuctEG
        }
        private static ElementType GetElementType(Element element)
        {
            int categoryId = element.Category.Id.IntegerValue;
            string typeParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString();
            string familyParam = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString();

            if (categoryId == (int)BuiltInCategory.OST_CableTray && typeParam?.Contains("tray") == true)
                return ElementType.CableTray;
            if (categoryId == (int)BuiltInCategory.OST_CableTrayFitting && familyParam?.Contains("bend") == true)
                return ElementType.CableTrayFitting;
            if (categoryId == (int)BuiltInCategory.OST_DuctCurves && typeParam?.Contains("ASML_ОГК") == true)
                return ElementType.DuctOGK;
            if (categoryId == (int)BuiltInCategory.OST_DuctCurves && typeParam?.Contains("ASML_ЭГ") == true)
                return ElementType.DuctEG;

            return ElementType.None;
        }
        private static bool IsValidElement(Element element)
        {
            return GetElementType(element) != ElementType.None;
        }
        private static void ProcessAllElements(IReadOnlyCollection<Element> elements,
            List<Tuple<string, ElementId>> failedElements)
        {
            var cableTrayMappings = new List<Tuple<string, string>>
            {
                Tuple.Create("DKC_Единица измерения", "ASML_Единица измерения"),
                Tuple.Create("DKC_Завод-изготовитель", "ASML_Завод-изготовитель"),
                Tuple.Create("DKC_Код изделия", "ASML_Код изделия"),
                Tuple.Create("DKC_Масса_Текст", "ASML_Масса_Текст"),
                Tuple.Create("DKC_Наименование", "ASML_Наименование"),
                Tuple.Create("И_НаименованиеСистемы", "ASML_Тип")
            };
            var ductMappings = new List<Tuple<string, string>>
            {
                Tuple.Create("ОГК_Наименование", "ASML_Наименование"),
                Tuple.Create("ОГК_Тип", "ASML_Тип"),
                Tuple.Create("ОГК_Завод-изготовитель", "ASML_Завод-изготовитель"),
                Tuple.Create("ОГК_Раздел спецификации", "ASML_Раздел спецификации"),
                Tuple.Create("ОГК_Единицы измерения", "ASML_Единица измерения")
            };
            var mzMappings = new List<Tuple<string, string>>
            {
                Tuple.Create("М_Наименование", "ASML_Наименование"),
                Tuple.Create("М_Масса", "ASML_Масса_Текст"),
                Tuple.Create("М_Раздел", "ASML_Раздел спецификации"),
                Tuple.Create("М_Единицы измерения", "ASML_Единица измерения")
            };
            foreach (var element in elements)
            {
                try
                {
                    string typeParam = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString();
                    SetQuantityParameter(element, failedElements, typeParam);
                    ElementType elementType = GetElementType(element);

                    switch (elementType)
                    {
                        case ElementType.CableTray:
                        case ElementType.CableTrayFitting:
                            CopyMappedParameters(element, cableTrayMappings, failedElements);
                            break;
                        case ElementType.DuctOGK:
                            CopyMappedParameters(element, ductMappings, failedElements);
                            break;
                        case ElementType.DuctEG:
                            CopyMappedParameters(element, mzMappings, failedElements);
                            break;
                        case ElementType.None:
                            AddFailedElement(element, failedElements, "Элемент не прошел фильтр");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    AddFailedElement(element, failedElements, ex.Message);
                }
            }
        }
        private static void SetQuantityParameter(Element element,
        List<Tuple<string, ElementId>> failedElements,
        string typeParam)
        {
            var qtyParam = element.LookupParameter("ASML_Количество");
            if (qtyParam == null)
            {
                AddFailedElement(element, failedElements, "Отсутствует параметр ASML_Количество");
                return;
            }

            int categoryId = element.Category.Id.IntegerValue;
            if (categoryId == (int)BuiltInCategory.OST_CableTray || categoryId == (int)BuiltInCategory.OST_DuctCurves)
            {
                var lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0.0;
                var widthParam = element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)?.AsDouble() ?? 0.0;

                double lengthInMeters = lengthParam * FEET_TO_METERS;
                double widthInMeters = widthParam * FEET_TO_METERS;

                double quantity = lengthInMeters;
                if (typeParam?.Contains("ASML_ОГК") == true)
                {
                    quantity = lengthInMeters * widthInMeters;
                }

                qtyParam.Set((double)Math.Round(quantity, 2));
            }
            else if (categoryId == (int)BuiltInCategory.OST_CableTrayFitting)
            {
                qtyParam.Set(1);
            }
        }
        private static void CopyMappedParameters(Element element,
            IEnumerable<Tuple<string, string>> paramMappings,
            List<Tuple<string, ElementId>> failedElements)
        {
            foreach (var mapping in paramMappings)
            {
                var sourceParam = element.LookupParameter(mapping.Item1);
                var targetParam = element.LookupParameter(mapping.Item2);

                if (sourceParam == null)
                {
                    AddFailedElement(element, failedElements, $"Источник {mapping.Item1} не найден");
                    continue;
                }
                if (targetParam == null)
                {
                    AddFailedElement(element, failedElements, $"Цель {mapping.Item2} не найдена");
                    continue;
                }
                if (targetParam.IsReadOnly)
                {
                    AddFailedElement(element, failedElements, $"Цель {mapping.Item2} только для чтения");
                    continue;
                }
                CopyParameterValue(sourceParam, targetParam);
            }
        }
        private static void CopyParameterValue(Parameter source, Parameter target)
        {
            switch (source.StorageType)
            {
                case StorageType.String: target.Set(source.AsString() ?? string.Empty); break;
                case StorageType.Double: target.Set(source.AsDouble()); break;
                case StorageType.Integer: target.Set(source.AsInteger()); break;
                case StorageType.ElementId: target.Set(source.AsElementId()); break;
            }
        }
        private static void AddFailedElement(Element element,
            List<Tuple<string, ElementId>> failedElements, string error)
        {
            string familyName = element.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString();
            failedElements.Add(Tuple.Create($"{familyName}: {error}", element.Id));
        }
        private static Result DisplayResults(IReadOnlyCollection<Tuple<string, ElementId>> failedElements,
            UIDocument uidoc, TimeSpan elapsed)
        {
            string message = failedElements.Any()
                ? $"Ошибки у семейств:\n{string.Join("\n", failedElements.Select(f => $"{f.Item1} (ID: {f.Item2})"))}"
                : $"Все элементы обработаны успешно.\nВремя обработки: {elapsed.TotalSeconds:F2} секунд";

            TaskDialog.Show(failedElements.Any() ? "Ошибки" : "Результат", message);
            return Result.Succeeded;
        }
    }
}
