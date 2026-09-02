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
        private ExpitVolumeResult _result;

        internal ExpitVolume(UIDocument uidoc)
        {
            _uidoc = uidoc ?? throw new ArgumentNullException(nameof(uidoc));
            InitializeComponent();

            _handler = new ExpitVolumeExternalEventHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            Closed += (sender, args) => _externalEvent.Dispose();
        }

        private Document Doc => _uidoc.Document;

        private void PickToposolid_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.PickToposolid);

        private void PickCutters_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.PickCutters);

        private void Calculate_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.Calculate);

        private void Write_Click(object sender, RoutedEventArgs e) =>
            Request(ExpitVolumeAction.Write);

        private void Close_Click(object sender, RoutedEventArgs e) => Close();

        private void Request(ExpitVolumeAction action)
        {
            _handler.Action = action;
            IsEnabled = false;
            StatusText.Foreground = Brushes.DarkSlateGray;
            StatusText.Text = action == ExpitVolumeAction.Calculate
                ? "Выполняется расчёт..."
                : action == ExpitVolumeAction.Write
                    ? "Выполняется запись..."
                    : string.Empty;
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
                    case ExpitVolumeAction.Calculate:
                        Calculate();
                        break;
                    case ExpitVolumeAction.Write:
                        Write();
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

        private void Calculate()
        {
            Element toposolid = GetToposolid();
            List<Element> cutters = GetCutters();
            _result = ExpitVolumeService.Calculate(Doc, toposolid, cutters);

            OriginalVolumeText.Text = FormatVolume(_result.OriginalCubicMeters);
            CutVolumeText.Text = FormatVolume(_result.CutCubicMeters);
            ExcavationVolumeText.Text = FormatVolume(_result.ExcavationCubicMeters);
            StatusText.Foreground = Brushes.DarkGreen;
            StatusText.Text = "Расчёт выполнен. Исходная геометрия восстановлена.";
            WriteButton.IsEnabled = true;
        }

        private void Write()
        {
            if (_result == null)
                throw new InvalidOperationException("Сначала выполните расчёт.");

            Element toposolid = GetToposolid();
            ExpitVolumeService.WriteResult(
                Doc,
                new List<Element> { toposolid },
                _result);

            StatusText.Foreground = Brushes.DarkGreen;
            StatusText.Text = "Значение записано в основное топотело.";
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
                throw new InvalidOperationException("Выберите хотя бы один элемент модели котлована.");
            return elements;
        }

        private void ResetResult()
        {
            _result = null;
            OriginalVolumeText.Text = "—";
            CutVolumeText.Text = "—";
            ExcavationVolumeText.Text = "—";
            StatusText.Text = string.Empty;
            WriteButton.IsEnabled = false;
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
        Calculate,
        Write
    }

    internal sealed class ExpitVolumeExternalEventHandler : IExternalEventHandler
    {
        private readonly ExpitVolume _form;
        internal ExpitVolumeAction Action { get; set; }

        internal ExpitVolumeExternalEventHandler(ExpitVolume form)
        {
            _form = form;
        }

        public void Execute(UIApplication app) => _form.ExecuteRequestedAction();

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

    internal sealed class ExpitVolumeResult
    {
        internal double OriginalInternal { get; }
        internal double CutInternal { get; }
        internal double ExcavationInternal => OriginalInternal - CutInternal;
        internal double OriginalCubicMeters => ToCubicMeters(OriginalInternal);
        internal double CutCubicMeters => ToCubicMeters(CutInternal);
        internal double ExcavationCubicMeters => ToCubicMeters(ExcavationInternal);

        internal ExpitVolumeResult(double originalInternal, double cutInternal)
        {
            OriginalInternal = originalInternal;
            CutInternal = cutInternal;
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

    internal static class ExpitVolumeService
    {
        internal const string VolumeParameterName = "КП_Р_Объем";
        private const double Tolerance = 1.0e-9;

        internal static ExpitVolumeResult Calculate(
            Document doc,
            Element toposolid,
            IList<Element> cutters)
        {
            Validate(doc, toposolid, cutters);

            foreach (Element cutter in cutters)
            {
                bool firstCutsSecond;
                if (SolidSolidCutUtils.CutExistsBetweenElements(
                    toposolid,
                    cutter,
                    out firstCutsSecond))
                    throw new InvalidOperationException(
                        $"Между топотелом и элементом ID {IDHelper.ElIdValue(cutter.Id)} " +
                        "уже существует Cut Geometry.");
            }

            double originalVolume = GetVolume(toposolid);
            double cutVolume;

            using (Transaction temporary = new Transaction(doc, "КП: временно вырезать котлован"))
            {
                if (temporary.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать временную транзакцию.");

                SetFailureHandling(temporary);
                try
                {
                    foreach (Element cutter in cutters)
                        SolidSolidCutUtils.AddCutBetweenSolids(doc, toposolid, cutter);

                    doc.Regenerate();
                    cutVolume = GetVolume(toposolid);
                }
                finally
                {
                    if (temporary.GetStatus() == TransactionStatus.Started)
                        temporary.RollBack();
                }
            }

            if (originalVolume + Tolerance < cutVolume)
                throw new InvalidOperationException("V2 оказался больше V1. Расчёт остановлен.");
            if (originalVolume - cutVolume <= Tolerance)
                throw new InvalidOperationException(
                    "Разность объёмов равна нулю. Проверьте пересечение элементов.");

            return new ExpitVolumeResult(originalVolume, cutVolume);
        }

        internal static void WriteResult(
            Document doc,
            IList<Element> targets,
            ExpitVolumeResult result)
        {
            if (targets == null || targets.Count == 0)
                throw new InvalidOperationException("Не выбраны элементы для записи.");

            List<Parameter> parameters = targets
                .Select(GetWritableVolumeParameter)
                .ToList();

            using (Transaction transaction = new Transaction(doc, "КП: записать объем котлована"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    throw new InvalidOperationException("Не удалось начать транзакцию записи.");

                SetFailureHandling(transaction);
                foreach (Parameter parameter in parameters)
                {
                    if (!parameter.Set(result.ExcavationInternal))
                    {
                        transaction.RollBack();
                        throw new InvalidOperationException("Revit не принял значение параметра.");
                    }
                }

                if (transaction.Commit() != TransactionStatus.Committed)
                    throw new InvalidOperationException("Не удалось записать объем котлована.");
            }
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
                    throw new InvalidOperationException("Топотело нельзя выбрать как модель котлована.");
            }
        }

        private static Parameter GetWritableVolumeParameter(Element element)
        {
            Parameter parameter = element?.LookupParameter(VolumeParameterName);
            if (parameter == null)
                throw new InvalidOperationException(
                    $"У элемента ID {IDHelper.ElIdValue(element.Id)} отсутствует параметр " +
                    $"«{VolumeParameterName}».");
            if (parameter.IsReadOnly)
                throw new InvalidOperationException(
                    $"Параметр «{VolumeParameterName}» доступен только для чтения.");
            if (parameter.StorageType != StorageType.Double || !IsVolumeParameter(parameter))
                throw new InvalidOperationException(
                    $"Параметр «{VolumeParameterName}» должен иметь тип данных «Объем».");
            return parameter;
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
                throw new InvalidOperationException("В геометрии топотела не найден объемный Solid.");
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
