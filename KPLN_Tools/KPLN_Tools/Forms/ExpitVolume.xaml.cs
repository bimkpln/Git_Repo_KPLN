using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_Tools.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace KPLN_Tools.Forms
{
    internal partial class ExpitVolume : Window
    {
        private readonly UIDocument _uidoc;
        private readonly ExpitVolumeExternalEventHandler _handler;
        private readonly ExternalEvent _externalEvent;
        private ElementId _toposolidId = ElementId.InvalidElementId;
        private readonly List<ElementId> _cutterIds = new List<ElementId>();
        private ExpitVolumeOperationResult _operation;

        internal ExpitVolume(UIDocument uidoc)
        {
            _uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));
            InitializeComponent();

            _handler = new ExpitVolumeExternalEventHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            Closed += (sender, args) => _externalEvent.Dispose();
            SetOperationState(false);
        }

        private Document Doc => _uidoc.Document;

        private void PickToposolid_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.PickToposolid);

        private void PickCutters_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.PickCutters);

        private void CalculateWrite_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.CalculateAndWrite);

        private void CancelOperation_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.CancelOperation);

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void Request(ExpitVolumeAction action)
        {
            _handler.Action = action;
            IsEnabled = false;
            StatusText.Foreground = Brushes.DarkSlateGray;

            if (action == ExpitVolumeAction.CalculateAndWrite)
                StatusText.Text = "Выполняется расчёт и запись...";
            else if (action == ExpitVolumeAction.CancelOperation)
                StatusText.Text = "Выполняется отмена...";
            else
                StatusText.Text = string.Empty;

            _externalEvent.Raise();
        }

        internal void ExecuteRequestedAction()
        {
            bool selectionAction =
                _handler.Action == ExpitVolumeAction.PickToposolid ||
                _handler.Action == ExpitVolumeAction.PickCutters;

            try
            {
                if (selectionAction)
                    Hide();

                switch (_handler.Action)
                {
                    case ExpitVolumeAction.PickToposolid:
                        PickToposolid();
                        break;
                    case ExpitVolumeAction.PickCutters:
                        PickCutters();
                        break;
                    case ExpitVolumeAction.CalculateAndWrite:
                        CalculateAndWrite();
                        break;
                    case ExpitVolumeAction.CancelOperation:
                        CancelOperation();
                        break;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                StatusText.Text = "Выбор отменён.";
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
            finally
            {
                _handler.Action = ExpitVolumeAction.None;
                if (IsLoaded)
                {
                    IsEnabled = true;
                    if (selectionAction && !IsVisible)
                        Show();
                    Activate();
                }
            }
        }

        private void PickToposolid()
        {
            Reference reference = _uidoc.Selection.PickObject(
                ObjectType.Element,
                new ToposolidSelectionFilter(),
                "Выберите топотело");

            _toposolidId = reference.ElementId;
            ToposolidText.Text = DescribeElement(Doc.GetElement(_toposolidId));
            ResetResult();
        }

        private void PickCutters()
        {
            IList<Reference> references = _uidoc.Selection.PickObjects(
                ObjectType.Element,
                new SolidElementSelectionFilter(),
                "Выберите элементы модели котлована и нажмите «Готово»");

            if (references.Count == 0)
                throw new InvalidOperationException("Не выбраны элементы модели котлована.");

            _cutterIds.Clear();
            _cutterIds.AddRange(references.Select(x => x.ElementId));

            List<Element> cutters = GetCutters();
            if (cutters.Count == 1)
            {
                CuttersText.Text = DescribeElement(cutters[0]);
                CuttersList.ItemsSource = null;
                CuttersList.Visibility = System.Windows.Visibility.Collapsed;
            }
            else
            {
                CuttersText.Text = $"Выбрано элементов: {cutters.Count}";
                CuttersList.ItemsSource = cutters.Select(DescribeElement).ToList();
                CuttersList.Visibility = System.Windows.Visibility.Visible;
            }

            ResetResult();
        }

        private void CalculateAndWrite()
        {
            Element toposolid = GetToposolid();
            List<Element> cutters = GetCutters();

            bool writeToCutters = WriteToCuttersRadioButton.IsChecked == true;
            _operation = ExpitVolumeService.CalculateAndWrite(
                Doc,
                toposolid,
                cutters,
                writeToCutters);
            _uidoc.RefreshActiveView();

            OriginalVolumeText.Text = FormatVolume(_operation.OriginalCubicMeters);
            CutVolumeText.Text = FormatVolume(_operation.CutCubicMeters);
            ExcavationVolumeText.Text = FormatVolume(_operation.ExcavationCubicMeters);
            StatusText.Foreground = Brushes.DarkGreen;
            StatusText.Text = writeToCutters
                ? "Вырез создан, значения записаны в элементы модели котлована."
                : "Вырез создан, значение записано в основное топотело.";
            SetOperationState(true);
        }

        private void CancelOperation()
        {
            if (_operation == null)
                throw new InvalidOperationException("Нет операции для отмены.");

            Element toposolid = GetToposolid();
            ExpitVolumeService.CancelOperation(Doc, toposolid, _operation);
            _uidoc.RefreshActiveView();

            _operation = null;
            OriginalVolumeText.Text = "—";
            CutVolumeText.Text = "—";
            ExcavationVolumeText.Text = "—";
            StatusText.Foreground = Brushes.DarkGreen;
            StatusText.Text = "Вырез и запись параметра отменены.";
            SetOperationState(false);
        }

        private void SetOperationState(bool operationApplied)
        {
            PickToposolidButton.IsEnabled = !operationApplied;
            PickCuttersButton.IsEnabled = !operationApplied;
            WriteTargetGroup.IsEnabled = !operationApplied;
            CalculateWriteButton.IsEnabled = !operationApplied;
            CancelOperationButton.IsEnabled = operationApplied;
        }

        private Element GetToposolid()
        {
            if (_toposolidId == ElementId.InvalidElementId)
                throw new InvalidOperationException("Выберите топотело.");

            Element element = Doc.GetElement(_toposolidId);
            if (element == null || !element.IsValidObject)
                throw new InvalidOperationException("Выбранное топотело недоступно.");
            return element;
        }

        private List<Element> GetCutters()
        {
            List<Element> elements = _cutterIds
                .Select(Doc.GetElement)
                .Where(x => x != null && x.IsValidObject)
                .ToList();

            if (elements.Count == 0)
                throw new InvalidOperationException(
                    "Выберите хотя бы один элемент модели котлована.");
            return elements;
        }

        private void ResetResult()
        {
            _operation = null;
            OriginalVolumeText.Text = "—";
            CutVolumeText.Text = "—";
            ExcavationVolumeText.Text = "—";
            StatusText.Text = string.Empty;
            SetOperationState(false);
        }

        private void ShowError(string text)
        {
            StatusText.Foreground = Brushes.Firebrick;
            StatusText.Text = text;
            TaskDialog.Show("Получение объема котлована", text);
        }

        private static string DescribeElement(Element element)
        {
            if (element == null)
                return "Элемент недоступен";

            string category = element.Category?.Name ?? "Без категории";
            return $"{category}: {element.Name} (ID {IDHelper.ElIdValue(element.Id)})";
        }

        private static string FormatVolume(double value) => $"{value:N2} м³";
    }

    internal enum ExpitVolumeAction
    {
        None,
        PickToposolid,
        PickCutters,
        CalculateAndWrite,
        CancelOperation
    }

    internal sealed class ExpitVolumeExternalEventHandler : IExternalEventHandler
    {
        private readonly ExpitVolume _form;
        internal ExpitVolumeAction Action { get; set; }

        internal ExpitVolumeExternalEventHandler(ExpitVolume form)
        {
            _form = form;
        }

        public void Execute(UIApplication app)
        {
            if (_form.IsLoaded)
                _form.ExecuteRequestedAction();
        }

        public string GetName() => "Получение объема котлована";
    }

    internal sealed class ToposolidSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element) =>
            element?.GetType().FullName == "Autodesk.Revit.DB.Toposolid";

        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    internal sealed class SolidElementSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element element)
        {
            if (element == null || element is ElementType)
                return false;

            try
            {
                return element.get_Geometry(new Options
                {
                    IncludeNonVisibleObjects = true,
                    DetailLevel = ViewDetailLevel.Fine
                }) != null;
            }
            catch
            {
                return false;
            }
        }

        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    internal sealed class ExpitVolumeOperationResult
    {
        internal double OriginalInternal { get; }
        internal double CutInternal { get; }
        internal double ExcavationInternal => OriginalInternal - CutInternal;
        internal double OriginalCubicMeters => ToCubicMeters(OriginalInternal);
        internal double CutCubicMeters => ToCubicMeters(CutInternal);
        internal double ExcavationCubicMeters => ToCubicMeters(ExcavationInternal);
        internal IReadOnlyList<ElementId> OperationCutterIds { get; }
        internal IReadOnlyList<ExpitVolumeParameterBackup> ParameterBackups { get; }

        internal ExpitVolumeOperationResult(
            double originalInternal,
            double cutInternal,
            IEnumerable<ElementId> operationCutterIds,
            IEnumerable<ExpitVolumeParameterBackup> parameterBackups)
        {
            OriginalInternal = originalInternal;
            CutInternal = cutInternal;
            OperationCutterIds = operationCutterIds.ToList().AsReadOnly();
            ParameterBackups = parameterBackups.ToList().AsReadOnly();
        }

        private static double ToCubicMeters(double value)
        {
#if Debug2020 || Revit2020
            return UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_CUBIC_METERS);
#else
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.CubicMeters);
#endif
        }
    }

    internal sealed class ExpitVolumeParameterBackup
    {
        internal ElementId ElementId { get; }
        internal bool HadValue { get; }
        internal double Value { get; }

        internal ExpitVolumeParameterBackup(
            ElementId elementId,
            bool hadValue,
            double value)
        {
            ElementId = elementId;
            HadValue = hadValue;
            Value = value;
        }
    }

    internal static class ExpitVolumeService
    {
        internal const string VolumeParameterName = "КП_Р_Объем";
        private const double Tolerance = 1.0e-9;

        internal static ExpitVolumeOperationResult CalculateAndWrite(
            Document doc,
            Element toposolid,
            IList<Element> cutters,
            bool writeToCutters)
        {
            Validate(doc, toposolid, cutters);
            ValidateCuttersForExcavation(toposolid, cutters);

            List<Element> parameterTargets = writeToCutters
                ? cutters.ToList()
                : new List<Element> { toposolid };
            Dictionary<ElementId, Parameter> parameters = parameterTargets
                .ToDictionary(x => x.Id, GetWritableVolumeParameter);
            List<ExpitVolumeParameterBackup> parameterBackups = parameterTargets
                .Select(x => CreateParameterBackup(x, parameters[x.Id]))
                .ToList();

            List<ElementId> operationCutterIds = cutters
                .Select(cutter => cutter.Id)
                .ToList();
            List<ElementId> cuttersToCreate = cutters
                .Where(cutter => !CutExists(toposolid, cutter))
                .Select(cutter => cutter.Id)
                .ToList();

            double originalVolume;
            double cutVolume;

            using (Transaction transaction = new Transaction(
                doc,
                "КП: рассчитать и записать объем котлована"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать транзакцию.");

                SetFailureHandling(transaction);
                try
                {
                    // Сначала создаём вырез и считываем V2.
                    foreach (Element cutter in cutters.Where(
                        x => cuttersToCreate.Contains(x.Id)))
                        CreateExcavation(doc, toposolid, cutter);

                    doc.Regenerate();

                    foreach (Element cutter in cutters)
                        VerifyCutterCutsToposolid(toposolid, cutter);

                    cutVolume = GetVolume(toposolid);

                    // Временно удаляем вырезы, считываем V1 и откатываем только
                    // это удаление. После RollBack вырезы снова остаются в модели.
                    using (SubTransaction temporaryRemoval = new SubTransaction(doc))
                    {
                        if (temporaryRemoval.Start() != TransactionStatus.Started)
                            throw new InvalidOperationException(
                                "Не удалось начать временную отмену вырезания.");

                        try
                        {
                            foreach (Element cutter in cutters)
                                RemoveExcavation(doc, toposolid, cutter);

                            doc.Regenerate();
                            originalVolume = GetVolume(toposolid);
                        }
                        finally
                        {
                            if (temporaryRemoval.GetStatus() == TransactionStatus.Started)
                                temporaryRemoval.RollBack();
                        }
                    }

                    doc.Regenerate();

                    foreach (Element cutter in cutters)
                        VerifyCutterCutsToposolid(toposolid, cutter);

                    Dictionary<ElementId, double> individualVolumes =
                        writeToCutters
                            ? CalculateIndividualVolumes(doc, toposolid, cutters, originalVolume)
                            : null;

                    ValidateVolumeDifference(originalVolume, cutVolume);

                    if (writeToCutters)
                    {
                        foreach (Element cutter in cutters)
                            SetVolumeParameter(
                                parameters[cutter.Id],
                                individualVolumes[cutter.Id],
                                cutter);
                    }
                    else
                    {
                        SetVolumeParameter(
                            parameters[toposolid.Id],
                            originalVolume - cutVolume,
                            toposolid);
                    }

                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException(
                            "Не удалось завершить расчёт и запись параметра.");
                }
                catch
                {
                    if (transaction.GetStatus() == TransactionStatus.Started)
                        transaction.RollBack();
                    throw;
                }
            }

            return new ExpitVolumeOperationResult(
                originalVolume,
                cutVolume,
                operationCutterIds,
                parameterBackups);
        }

        internal static void CancelOperation(
            Document doc,
            Element toposolid,
            ExpitVolumeOperationResult operation)
        {
            if (operation == null)
                throw new ArgumentNullException(nameof(operation));

            using (Transaction transaction = new Transaction(
                doc,
                "КП: отменить расчет объема котлована"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать транзакцию отмены.");

                SetFailureHandling(transaction);
                try
                {
                    foreach (ElementId cutterId in operation.OperationCutterIds)
                    {
                        Element cutter = doc.GetElement(cutterId);
                        if (cutter != null)
                            RemoveExcavation(doc, toposolid, cutter);
                    }

                    doc.Regenerate();

                    foreach (ExpitVolumeParameterBackup backup in operation.ParameterBackups)
                        RestoreParameter(doc, backup);

                    if (transaction.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException(
                            "Не удалось отменить вырез и запись параметра.");
                }
                catch
                {
                    if (transaction.GetStatus() == TransactionStatus.Started)
                        transaction.RollBack();
                    throw;
                }
            }
        }

        private static bool CutExists(Element first, Element second)
        {
            bool firstCutsSecond;
            return TryGetCutOrder(first, second, out firstCutsSecond);
        }

        private static bool TryGetCutOrder(
            Element first,
            Element second,
            out bool firstCutsSecond)
        {
            return SolidSolidCutUtils.CutExistsBetweenElements(
                first, second, out firstCutsSecond);
        }

        private static void CreateExcavation(
            Document doc,
            Element toposolid,
            Element cutter)
        {
            SolidSolidCutUtils.AddCutBetweenSolids(doc, toposolid, cutter);
        }

        private static void VerifyCutterCutsToposolid(
            Element toposolid,
            Element cutter)
        {
            bool firstCutsSecond;
            if (!TryGetCutOrder(toposolid, cutter, out firstCutsSecond))
                throw new InvalidOperationException(
                    $"Не удалось создать вырез элементом ID " +
                    $"{IDHelper.ElIdValue(cutter.Id)}.");

            if (firstCutsSecond)
                throw new InvalidOperationException(
                    $"Revit создал вырез в обратном направлении для элемента ID " +
                    $"{IDHelper.ElIdValue(cutter.Id)}.");
        }

        private static void RemoveExcavation(
            Document doc,
            Element toposolid,
            Element cutter)
        {
            if (CutExists(toposolid, cutter))
                SolidSolidCutUtils.RemoveCutBetweenSolids(doc, toposolid, cutter);
        }

        private static void ValidateCuttersForExcavation(
            Element toposolid,
            IList<Element> cutters)
        {
            foreach (Element cutter in cutters)
            {
                if (CutExists(toposolid, cutter))
                {
                    VerifyCutterCutsToposolid(toposolid, cutter);
                    continue;
                }

                CutFailureReason reason;
                if (!SolidSolidCutUtils.CanElementCutElement(cutter, toposolid, out reason))
                    throw new InvalidOperationException(
                        $"Элемент ID {IDHelper.ElIdValue(cutter.Id)} не может вырезать " +
                        $"топотело: {reason}.");
            }
        }

        private static Dictionary<ElementId, double> CalculateIndividualVolumes(
            Document doc,
            Element toposolid,
            IList<Element> cutters,
            double originalVolume)
        {
            Dictionary<ElementId, double> result = new Dictionary<ElementId, double>();

            foreach (Element current in cutters)
            {
                double volumeWithCurrent;
                using (SubTransaction isolateCurrent = new SubTransaction(doc))
                {
                    if (isolateCurrent.Start() != TransactionStatus.Started)
                        throw new InvalidOperationException(
                            "Не удалось начать отдельный расчёт элемента котлована.");

                    try
                    {
                        foreach (Element other in cutters.Where(x => x.Id != current.Id))
                            RemoveExcavation(doc, toposolid, other);

                        doc.Regenerate();
                        volumeWithCurrent = GetVolume(toposolid);
                    }
                    finally
                    {
                        if (isolateCurrent.GetStatus() == TransactionStatus.Started)
                            isolateCurrent.RollBack();
                    }
                }

                doc.Regenerate();
                ValidateVolumeDifference(originalVolume, volumeWithCurrent);
                result[current.Id] = originalVolume - volumeWithCurrent;
            }

            foreach (Element cutter in cutters)
                VerifyCutterCutsToposolid(toposolid, cutter);

            return result;
        }

        private static void ValidateVolumeDifference(double originalVolume, double cutVolume)
        {
            if (originalVolume + Tolerance < cutVolume)
                throw new InvalidOperationException("V2 оказался больше V1. Расчёт остановлен.");
            if (originalVolume - cutVolume <= Tolerance)
                throw new InvalidOperationException(
                    "Разность объёмов равна нулю. Проверьте пересечение элементов.");
        }

        private static void Validate(Document doc, Element toposolid, IList<Element> cutters)
        {
            if (doc == null)
                throw new ArgumentNullException(nameof(doc));
            if (toposolid == null || !toposolid.IsValidObject)
                throw new InvalidOperationException("Топотело не выбрано или удалено.");
            if (toposolid.GetType().FullName != "Autodesk.Revit.DB.Toposolid")
                throw new InvalidOperationException("Первый элемент не является топотелом.");
            if (cutters == null || cutters.Count == 0)
                throw new InvalidOperationException("Не выбраны элементы модели котлована.");

            foreach (Element cutter in cutters)
            {
                if (cutter == null || !cutter.IsValidObject)
                    throw new InvalidOperationException("Один из элементов котлована недоступен.");
                if (cutter.Id == toposolid.Id)
                    throw new InvalidOperationException(
                        "Топотело нельзя выбрать как модель котлована.");
            }
        }

        private static Parameter GetWritableVolumeParameter(Element element)
        {
            Parameter parameter = element?.LookupParameter(VolumeParameterName);
            if (parameter == null)
                throw new InvalidOperationException(
                    $"У элемента ID {IDHelper.ElIdValue(element.Id)} отсутствует " +
                    $"параметр «{VolumeParameterName}».");
            if (parameter.IsReadOnly)
                throw new InvalidOperationException(
                    $"Параметр «{VolumeParameterName}» элемента ID " +
                    $"{IDHelper.ElIdValue(element.Id)} доступен только для чтения.");
            if (parameter.StorageType != StorageType.Double || !IsVolumeParameter(parameter))
                throw new InvalidOperationException(
                    $"Параметр «{VolumeParameterName}» должен иметь тип данных «Объем».");
            return parameter;
        }

        private static ExpitVolumeParameterBackup CreateParameterBackup(
            Element element,
            Parameter parameter) =>
            new ExpitVolumeParameterBackup(
                element.Id,
                parameter.HasValue,
                parameter.HasValue ? parameter.AsDouble() : 0.0);

        private static void SetVolumeParameter(
            Parameter parameter,
            double value,
            Element element)
        {
            if (!parameter.Set(value))
                throw new InvalidOperationException(
                    $"Revit не принял значение параметра «{VolumeParameterName}» " +
                    $"для элемента ID {IDHelper.ElIdValue(element.Id)}.");
        }

        private static void RestoreParameter(
            Document doc,
            ExpitVolumeParameterBackup backup)
        {
            Element element = doc.GetElement(backup.ElementId);
            if (element == null)
                return;

            Parameter parameter = GetWritableVolumeParameter(element);
            if (backup.HadValue)
            {
                SetVolumeParameter(parameter, backup.Value, element);
            }
            else
            {
                // Revit 2024 запрещает ClearValue(), если у определения параметра
                // HideWhenNoValue == false. В таком случае безопасно возвращаем 0 м³.
                try
                {
                    if (!parameter.ClearValue())
                        SetVolumeParameter(parameter, 0.0, element);
                }
                catch (Exception)
                {
                    SetVolumeParameter(parameter, 0.0, element);
                }
            }
        }

        private static double GetVolume(Element element)
        {
            GeometryElement geometry = element.get_Geometry(new Options
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = true,
                DetailLevel = ViewDetailLevel.Fine
            });

            double volume = SumVolume(geometry);
            if (volume <= Tolerance)
                throw new InvalidOperationException(
                    "В геометрии топотела не найден объемный Solid.");
            return volume;
        }

        private static double SumVolume(GeometryElement geometry)
        {
            if (geometry == null)
                return 0.0;

            double result = 0.0;
            foreach (GeometryObject obj in geometry)
            {
                if (obj is Solid solid && solid.Faces.Size > 0 && solid.Edges.Size > 0)
                {
                    try
                    {
                        if (solid.Volume > Tolerance)
                            result += solid.Volume;
                    }
                    catch (Exception)
                    {
                    }
                }
                else if (obj is GeometryInstance instance)
                {
                    result += SumVolume(instance.GetInstanceGeometry());
                }
            }
            return result;
        }

        private static bool IsVolumeParameter(Parameter parameter)
        {
#if Debug2020 || Revit2020
            return parameter.Definition.ParameterType == ParameterType.Volume;
#else
            return parameter.Definition.GetDataType() == SpecTypeId.Volume;
#endif
        }

        private static void SetFailureHandling(Transaction transaction)
        {
            FailureHandlingOptions options = transaction.GetFailureHandlingOptions();
            options.SetFailuresPreprocessor(new SilentFailuresPreprocessor());
            options.SetClearAfterRollback(true);
            transaction.SetFailureHandlingOptions(options);
        }
    }

    internal sealed class SilentFailuresPreprocessor : IFailuresPreprocessor
    {
        public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
        {
            bool hasErrors = false;
            foreach (FailureMessageAccessor failure in failuresAccessor.GetFailureMessages())
            {
                if (failure.GetSeverity() == FailureSeverity.Warning)
                    failuresAccessor.DeleteWarning(failure);
                else
                    hasErrors = true;
            }

            return hasErrors
                ? FailureProcessingResult.ProceedWithRollBack
                : FailureProcessingResult.Continue;
        }
    }
}
