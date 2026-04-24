using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    public class Command_SET_EOMParams : IExternalCommand
    {
        #region Вспомогательные типы (объявлены первыми для видимости)
        private enum ElementType { None, CableTray, CableTrayFitting, DuctEG }
        private class FailedElementInfo
        {
            public string ElementName { get; }
            public ElementId Id { get; }
            public string Error { get; }

            public FailedElementInfo(string name, ElementId id, string error)
            {
                ElementName = name;
                Id = id;
                Error = error;
            }
        }

        #endregion

        private const double FEET_TO_METERS = 0.3048;
        private readonly Dictionary<string, Definition> _paramDefCache = new Dictionary<string, Definition>();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uidoc = commandData.Application.ActiveUIDocument;
            var doc = uidoc.Document;
            var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                                       string.Format("ASML_Log_{0}.txt", DateTime.Now.ToString("yyyyMMdd_HHmmss")));
            var selectedIds = uidoc.Selection.GetElementIds();
            var elementsToProcess = selectedIds.Count > 0
                ? GetElementsFromSelection(doc, selectedIds)
                : CollectAllTargetElements(doc);

            elementsToProcess = ExpandElementsWithNested(elementsToProcess, doc)
                .Where(e => e.Category != null && GetElementType(e) != ElementType.None)
                .ToList();

            if (elementsToProcess.Count == 0)
            {
                TaskDialog.Show("Предупреждение", "Подходящие элементы не найдены.");
                return Result.Succeeded;
            }

            SimpleProgressForm progressForm = new SimpleProgressForm(elementsToProcess.Count, "Обработка элементов...");
            StreamWriter logWriter = new StreamWriter(logPath, false, System.Text.Encoding.UTF8);

            try
            {
                progressForm.Show();
                logWriter.WriteLine(string.Format("=== Запуск обработки: {0} ===", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")));
                logWriter.WriteLine(string.Format("Всего элементов: {0}", elementsToProcess.Count));
                logWriter.WriteLine(string.Format("Режим: {0}", selectedIds.Count > 0 ? "Выделенные элементы" : "Весь документ"));
                logWriter.WriteLine(new string('-', 40));

                var stopwatch = Stopwatch.StartNew();
                var failedElements = new List<FailedElementInfo>();

                using (var tx = new Transaction(doc, "Заполнение параметров лотков и воздуховодов"))
                {
                    tx.Start();
                    try
                    {
                        ProcessElementsWithProgress(elementsToProcess, failedElements, progressForm, logWriter);
                        tx.Commit();
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        logWriter.WriteLine(string.Format("КРИТИЧЕСКАЯ ОШИБКА: {0}", ex.Message));
                        TaskDialog.Show("Критическая ошибка", string.Format("Не удалось записать значение:\n{0}", ex.Message));
                        return Result.Failed;
                    }
                }
                stopwatch.Stop();
                logWriter.WriteLine(new string('-', 40));
                logWriter.WriteLine(string.Format("Завершено: {0}", DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss")));
                logWriter.WriteLine(string.Format("Обработано успешно: {0}", elementsToProcess.Count - failedElements.Count));
                logWriter.WriteLine(string.Format("Ошибок: {0}", failedElements.Count));
                logWriter.WriteLine(string.Format("Время выполнения: {0:F2} сек.", stopwatch.Elapsed.TotalSeconds));
                return DisplayResults(failedElements, logPath);
            }
            finally
            {
                progressForm?.Close();
                logWriter?.Close();
            }


        }

        #region Сбор элементов
        private static List<Element> CollectAllTargetElements(Document doc)
        {
            var categoryFilters = new List<ElementFilter>
        {
            new ElementCategoryFilter(BuiltInCategory.OST_CableTray),
            new ElementCategoryFilter(BuiltInCategory.OST_CableTrayFitting),
            new ElementCategoryFilter(BuiltInCategory.OST_DuctCurves)
        };

            var baseElements = new FilteredElementCollector(doc)
                .WherePasses(new LogicalOrFilter(categoryFilters))
                .WhereElementIsNotElementType()
                .ToList();

            return ExpandElementsWithNested(baseElements, doc)
                .Where(e => e.Category != null && GetElementType(e) != ElementType.None)
                .ToList();
        }

        private static List<Element> GetElementsFromSelection(Document doc, ICollection<ElementId> selectedIds)
        {
            return selectedIds
                .Select(id => doc.GetElement(id))
                .Where(e => e != null && e.Category != null && GetElementType(e) != ElementType.None)
                .ToList();
        }
        #endregion

        #region Вложенные элементы

        private static List<Element> ExpandElementsWithNested(List<Element> source, Document doc)
        {
            var result = new List<Element>();
            var visited = new HashSet<ElementId>();

            foreach (var element in source)
            {
                CollectRecursive(element, doc, result, visited);
            }

            return result;
        }

        private static void CollectRecursive(Element element, Document doc, List<Element> result, HashSet<ElementId> visited)
        {
            if (element == null || visited.Contains(element.Id))
                return;

            visited.Add(element.Id);
            result.Add(element);

            if (element is FamilyInstance fi)
            {
                var subIds = fi.GetSubComponentIds();

                if (subIds == null || subIds.Count == 0)
                    return;

                foreach (var id in subIds)
                {
                    var sub = doc.GetElement(id);
                    CollectRecursive(sub, doc, result, visited);
                }
            }
        }

        #endregion

        #region Логика обработки
        private void ProcessElementsWithProgress(IReadOnlyList<Element> elements,
                                                 List<FailedElementInfo> failedElements,
                                                 SimpleProgressForm progressForm,
                                                 StreamWriter logWriter)
        {
            var cableTrayMappings = new List<Tuple<string, string>>
        {
            Tuple.Create("DKC_Единица измерения", "ASML_Единица измерения"),
            Tuple.Create("DKC_Завод-изготовитель", "ASML_Завод-изготовитель"),
            Tuple.Create("DKC_Код изделия", "ASML_Код изделия"),
            Tuple.Create("DKC_Масса_Текст", "ASML_Масса_Текст"),
            Tuple.Create("DKC_Наименование", "ASML_Наименование"),
            Tuple.Create("DKC_Обозначение", "ASML_Тип")
        };
            var mzMappings = new List<Tuple<string, string>>
        {
            Tuple.Create("М_Наименование", "ASML_Наименование"),
            Tuple.Create("М_Масса", "ASML_Масса_Текст"),
            Tuple.Create("М_Раздел", "ASML_Раздел спецификации"),
            Tuple.Create("М_Единицы измерения", "ASML_Единицы измерения")
        };

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                try
                {
                    var elementType = GetElementType(element);
                    string typeParam = GetParamValue(element, BuiltInParameter.ELEM_TYPE_PARAM);

                    SetQuantityParameter(element);

                    switch (elementType)
                    {
                        case ElementType.CableTray:
                        case ElementType.CableTrayFitting:
                            CopyMappedParameters(element, cableTrayMappings);
                            break;
                        case ElementType.DuctEG:
                            CopyMappedParameters(element, mzMappings);
                            break;
                        default:
                            throw new InvalidOperationException("Элемент не прошел внутренний фильтр");
                    }
                }
                catch (Exception ex)
                {
                    string familyName = GetParamValue(element, BuiltInParameter.ELEM_FAMILY_PARAM);
                    if (string.IsNullOrEmpty(familyName)) familyName = "Неизвестно";

                    failedElements.Add(new FailedElementInfo(familyName, element.Id, ex.Message));
                    logWriter.WriteLine(string.Format("ID: {0} | {1} | Ошибка: {2}", element.Id, familyName, ex.Message));
                }

                int progress = (int)((i + 1) * 100.0 / elements.Count);
                if (progress % 1 == 0 || (i + 1) % 50 == 0)
                {
                    progressForm.UpdateProgress(progress, string.Format("Обработка {0}/{1}", i + 1, elements.Count));
                    System.Windows.Forms.Application.DoEvents();
                }
            }
            progressForm.UpdateProgress(100, "Готово!");
        }

        private void SetQuantityParameter(Element element)
        {
            var qtyParam = GetParam(element, "ASML_Количество")
                ?? throw new InvalidOperationException("Нет ASML_Количество");

            int categoryId = element.Category.Id.IntegerValue;

            if (categoryId == (int)BuiltInCategory.OST_CableTray ||
                categoryId == (int)BuiltInCategory.OST_DuctCurves)
            {
                var lengthParam = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);

                double length = lengthParam?.AsDouble() ?? 0.0;
                double meters = length * FEET_TO_METERS;

                qtyParam.Set(Math.Round(meters, 2));
                return;
            }
            var source = GetParam(element, "DKC_Количество");

            if (source == null)
            {
                qtyParam.Set(1);
                return;
            }

            switch (source.StorageType)
            {
                case StorageType.Double:
                    qtyParam.Set(source.AsDouble());
                    break;

                case StorageType.Integer:
                    qtyParam.Set(source.AsInteger());
                    break;

                case StorageType.String:
                    if (double.TryParse(source.AsString(), out double val))
                        qtyParam.Set(val);
                    else
                        qtyParam.Set(1);
                    break;

                default:
                    qtyParam.Set(1);
                    break;
            }
        }

        private void CopyMappedParameters(Element element, IEnumerable<Tuple<string, string>> mappings)
        {
            foreach (var mapping in mappings)
            {
                var sourceParam = GetParam(element, mapping.Item1);
                if (sourceParam == null) throw new InvalidOperationException(string.Format("Источник {0} не найден", mapping.Item1));

                var targetParam = GetParam(element, mapping.Item2);
                if (targetParam == null) throw new InvalidOperationException(string.Format("Цель {0} не найдена", mapping.Item2));
                if (targetParam.IsReadOnly) throw new InvalidOperationException(string.Format("Цель {0} только для чтения", mapping.Item2));

                CopyParameterValue(sourceParam, targetParam);
            }
        }

        private static void CopyParameterValue(Parameter source, Parameter target)
        {
            if (source.StorageType != target.StorageType)
                throw new InvalidOperationException(string.Format("Несоответствие типов: {0} → {1}", source.StorageType, target.StorageType));

            switch (source.StorageType)
            {
                case StorageType.String:
                    target.Set(source.AsString() ?? string.Empty);
                    break;
                case StorageType.Double:
                    target.Set(source.AsDouble());
                    break;
                case StorageType.Integer:
                    target.Set(source.AsInteger());
                    break;
                case StorageType.ElementId:
                    target.Set(source.AsElementId());
                    break;
            }
        }
        #endregion

        #region Кэширование параметров
        private Parameter GetParam(Element element, string paramName)
        {
            if (!_paramDefCache.TryGetValue(paramName, out Definition def))
            {
                var p = element.LookupParameter(paramName);
                if (p != null)
                {
                    def = p.Definition;
                    _paramDefCache[paramName] = def;
                    return p;
                }
                return null;
            }
            return element.get_Parameter(def);
        }

        private static string GetParamValue(Element element, BuiltInParameter bip)
        {
            var param = element.get_Parameter(bip);
            return param != null ? param.AsValueString() : null;
        }
        #endregion

        #region Вспомогательные методы
        private static ElementType GetElementType(Element element)
        {
            int categoryId = element.Category.Id.IntegerValue;

            var typeElem = element.Document.GetElement(element.GetTypeId()) as Autodesk.Revit.DB.ElementType;
            string typeName = typeElem?.Name ?? "";

            if (categoryId == (int)BuiltInCategory.OST_CableTray)
                return ElementType.CableTray;

            if (categoryId == (int)BuiltInCategory.OST_CableTrayFitting)
                return ElementType.CableTrayFitting;

            if (categoryId == (int)BuiltInCategory.OST_DuctCurves &&
                typeName.IndexOf("ASML_ЭГ", StringComparison.OrdinalIgnoreCase) >= 0)
                return ElementType.DuctEG;

            return ElementType.None;
        }

        private static Result DisplayResults(IReadOnlyCollection<FailedElementInfo> failedElements, string logPath)
        {
            string title = failedElements.Any() ? "Завершено с ошибками" : "Результат";
            string msg = failedElements.Any()
                ? string.Format("Обработка завершена.\nОшибок: {0}\nЛог сохранён: {1}", failedElements.Count, logPath)
                : string.Format("Все элементы обработаны успешно.\nЛог сохранён: {0}", logPath);

            TaskDialog.Show(title, msg);
            return Result.Succeeded;
        }
        #endregion
    }
    #region UI: Прогресс-бар
    internal class SimpleProgressForm : System.Windows.Forms.Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _label;

        public SimpleProgressForm(int max, string title)
        {
            Text = title;
            Width = 400;
            Height = 120;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ControlBox = false;
            TopMost = true;

            _label = new Label { Left = 20, Top = 20, Width = 340, Text = "Подготовка..." };
            _progressBar = new ProgressBar { Left = 20, Top = 50, Width = 340, Maximum = max, Step = 1 };

            Controls.Add(_label);
            Controls.Add(_progressBar);
        }

        public void UpdateProgress(int value, string statusText)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateProgress(value, statusText)));
                return;
            }
            _progressBar.Value = Math.Min(value, _progressBar.Maximum);
            _label.Text = statusText;
        }
    }
    #endregion
}
