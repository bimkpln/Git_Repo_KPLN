using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_MEPSpacing.Common;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_MEPSpacing.Forms.Entities
{
    public enum SpacingCalculationMode
    {
        Centerline,
        Clearance
    }

    public enum UserMessageLevel
    {
        Neutral,
        Success,
        Warning,
        Error
    }

    /// <summary>
    /// Модель окна настройки равного шага MEP-элементов.
    /// </summary>
    public sealed class MepSpacingM : INotifyPropertyChanged
    {
        private string _distanceText = "100";
        private SpacingCalculationMode _calculationMode = SpacingCalculationMode.Centerline;
        private readonly List<ElementId> _selectedElementIds;
        private readonly List<ElementId> _baseElementIds = new List<ElementId>();
        private string _userMainStatus;
        private string _userHelp;
        private UserMessageLevel _messageLevel = UserMessageLevel.Neutral;

        public MepSpacingM(UIApplication uiapp, IEnumerable<ElementId> selectedElementIds)
        {
            UIApp = uiapp;
            Doc = uiapp.ActiveUIDocument.Document;
            _selectedElementIds = selectedElementIds == null
                ? new List<ElementId>()
                : selectedElementIds.ToList();

            SetStatus("Проверь расстояние и способ расчёта.", UserMessageLevel.Neutral);
            UserHelp = "Выбери элементы для расчёта. Если базовые элементы не выбраны, первый элемент в ряду останется на месте.";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public UIApplication UIApp { get; }

        public Document Doc { get; }

        public IReadOnlyList<ElementId> SelectedElementIds => _selectedElementIds;

        public IReadOnlyList<ElementId> BaseElementIds => _baseElementIds;

        public string DistanceText
        {
            get => _distanceText;
            set
            {
                _distanceText = value;
                NotifyPropertyChanged();
                NotifyStateChanged();
            }
        }

        public SpacingCalculationMode CalculationMode
        {
            get => _calculationMode;
            set
            {
                if (_calculationMode == value)
                    return;

                _calculationMode = value;
                NotifyPropertyChanged();
            }
        }

        public string UserMainStatus
        {
            get => _userMainStatus;
            set
            {
                _userMainStatus = string.IsNullOrWhiteSpace(value) ? string.Empty : $"ВАЖНО: {value}";
                NotifyPropertyChanged();
            }
        }

        public UserMessageLevel MessageLevel
        {
            get => _messageLevel;
            set
            {
                if (_messageLevel == value)
                    return;

                _messageLevel = value;
                NotifyPropertyChanged();
            }
        }

        public string UserHelp
        {
            get => _userHelp;
            set
            {
                _userHelp = value ?? string.Empty;
                NotifyPropertyChanged();
            }
        }

        public string SelectionInfo => $"Выбрано элементов для расчёта: {SelectedElementIds.Count}.";

        public string BaseSelectionInfo => _baseElementIds.Count == 0
            ? "Базовые элементы не выбраны. Отсчёт пойдёт от первого элемента в ряду."
            : $"Базовых элементов: {_baseElementIds.Count}.";

        public bool CanRun => GetAllElementIds().Count >= 2 && TryGetDistanceMm(out double distance) && distance > 0;

        public bool TryGetDistanceMm(out double distance)
        {
            string preparedValue = (DistanceText ?? string.Empty).Trim().Replace(',', '.');
            return double.TryParse(preparedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out distance);
        }

        public void SetResultStatus(int movedCount)
        {
            SetStatus($"Готово. Перемещено элементов: {movedCount}.", UserMessageLevel.Success);
        }

        public void SetResultStatus(string message)
        {
            SetStatus(message, UserMessageLevel.Success);
        }

        public void SetResultStatus(SpacingApplyResult result)
        {
            if (result == null)
                return;

            UserMessageLevel level = result.SkippedElementCount > 0 || result.Messages.Any()
                ? UserMessageLevel.Warning
                : UserMessageLevel.Success;

            SetStatus($"Готово. Перемещено элементов: {result.MovedElementCount}. Неподвижных элементов: {result.FixedElementCount}. Пропущено элементов: {result.SkippedElementCount}.", level);
            UserHelp = string.Join("\n", result.Messages.Take(4));
        }

        public void SetErrorStatus(string message)
        {
            SetStatus(message, UserMessageLevel.Error);
        }

        public void SetWarningStatus(string message)
        {
            SetStatus(message, UserMessageLevel.Warning);
        }

        public void SetBaseElementIds(IEnumerable<ElementId> elementIds)
        {
            _baseElementIds.Clear();

            if (elementIds != null)
                _baseElementIds.AddRange(elementIds.Where(id => id != null && !id.Equals(ElementId.InvalidElementId)));

            foreach (ElementId id in _baseElementIds)
            {
                if (!_selectedElementIds.Any(selectedId => GetElementIdValue(selectedId) == GetElementIdValue(id)))
                    _selectedElementIds.Add(id);
            }

            NotifyPropertyChanged(nameof(BaseElementIds));
            NotifyPropertyChanged(nameof(SelectionInfo));
            NotifyPropertyChanged(nameof(BaseSelectionInfo));
            NotifyStateChanged();
        }

        public void SetSelectedElementIds(IEnumerable<ElementId> elementIds)
        {
            _selectedElementIds.Clear();

            if (elementIds != null)
                _selectedElementIds.AddRange(elementIds.Where(id => id != null && !id.Equals(ElementId.InvalidElementId)));

            _baseElementIds.RemoveAll(baseId => !_selectedElementIds.Any(selectedId => GetElementIdValue(selectedId) == GetElementIdValue(baseId)));

            NotifyPropertyChanged(nameof(SelectedElementIds));
            NotifyPropertyChanged(nameof(BaseElementIds));
            NotifyPropertyChanged(nameof(SelectionInfo));
            NotifyPropertyChanged(nameof(BaseSelectionInfo));
            NotifyStateChanged();
        }

        public IReadOnlyList<ElementId> GetAllElementIds()
        {
            return _selectedElementIds
                .Concat(_baseElementIds)
                .Where(id => id != null && !id.Equals(ElementId.InvalidElementId))
                .GroupBy(GetElementIdValue)
                .Select(group => group.First())
                .ToList();
        }

        public void NotifyStateChanged()
        {
            NotifyPropertyChanged(nameof(CanRun));
        }

        private void SetStatus(string message, UserMessageLevel level)
        {
            UserMainStatus = message;
            MessageLevel = level;
        }

        private static long GetElementIdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
