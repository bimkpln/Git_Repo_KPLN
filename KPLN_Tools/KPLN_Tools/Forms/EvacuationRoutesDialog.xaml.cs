using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace KPLN_Tools.ExternalCommands.UI
{
    public sealed class EvacuationRoutesWorksetOption
    {
        public int Id { get; }
        public string Name { get; }

        public EvacuationRoutesWorksetOption(int id, string name)
        {
            Id = id;
            Name = name;
        }
    }

    public sealed class EvacuationRoutesStairListItem : INotifyPropertyChanged
    {
        private bool _isIncluded = true;
        private string _statusText;
        private Brush _statusBrush;
        private EvacuationRoutesStatus _status;

        public long ElementId { get; set; }
        public string Kind { get; set; }
        public string Name { get; set; }
        public string TypeName { get; set; }
        public string WorksetName { get; set; }
        public int RunCount { get; set; }
        public int LandingCount { get; set; }
        public int NestedCount { get; set; }
        public int ConnectedLevelCount { get; set; }
        public long? ParentMultistoryId { get; set; }
        public List<long> NestedStairIds { get; set; } = new List<long>();
        public bool HasExistingRoute { get; set; }
        public EvacuationRoutesStatus ExistingRouteStatus { get; set; } = EvacuationRoutesStatus.NotChecked;

        public string ElementsText
        {
            get
            {
                var parts = new List<string>();
                if (RunCount > 0)
                    parts.Add(FormatElementCount(RunCount, "Марш"));
                if (LandingCount > 0)
                    parts.Add(FormatElementCount(LandingCount, "Площадка"));

                return parts.Count == 0 ? "" : string.Join("; ", parts);
            }
        }

        public string NestedText
        {
            get
            {
                if (ConnectedLevelCount > 0)
                    return $"{ConnectedLevelCount} ур.";

                if (NestedCount > 0)
                    return $"{NestedCount} разм.";

                return ParentMultistoryId.HasValue ? $"в {ParentMultistoryId.Value}" : "";
            }
        }

        public Thickness ObjectIndent => ParentMultistoryId.HasValue ? new Thickness(22, 0, 0, 0) : new Thickness(0);
        public string ObjectText => ParentMultistoryId.HasValue ? "└ " + Kind : Kind;

        private static string FormatElementCount(int count, string name)
        {
            return count == 1 ? name : $"{count}x {name}";
        }

        public bool IsIncluded
        {
            get { return _isIncluded; }
            set
            {
                if (_isIncluded == value) return;
                _isIncluded = value;
                OnPropertyChanged(nameof(IsIncluded));
            }
        }

        public EvacuationRoutesStatus Status
        {
            get { return _status; }
            private set
            {
                if (_status == value) return;
                _status = value;
                OnPropertyChanged(nameof(Status));
            }
        }

        public string StatusText
        {
            get { return _statusText; }
            set
            {
                if (_statusText == value) return;
                _statusText = value;
                OnPropertyChanged(nameof(StatusText));
            }
        }

        public Brush StatusBrush
        {
            get { return _statusBrush; }
            set
            {
                if (_statusBrush == value) return;
                _statusBrush = value;
                OnPropertyChanged(nameof(StatusBrush));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public EvacuationRoutesStairListItem()
        {
            SetStatus(EvacuationRoutesStatus.NotChecked, "Не проверялось", null);
        }

        public void SetStatus(EvacuationRoutesStatus status, string text, string toolTip)
        {
            Status = status;
            StatusText = string.IsNullOrWhiteSpace(text) ? GetDefaultStatusText(status) : text;
            StatusBrush = GetStatusBrush(status);
        }

        private static string GetDefaultStatusText(EvacuationRoutesStatus status)
        {
            switch (status)
            {
                case EvacuationRoutesStatus.Ok:
                    return "ОК";
                case EvacuationRoutesStatus.Warning:
                    return "Проблемы";
                case EvacuationRoutesStatus.Error:
                    return "Не построено";
                case EvacuationRoutesStatus.Built:
                    return "Построено";
                case EvacuationRoutesStatus.PartialBuilt:
                    return "Частично";
                default:
                    return "Не проверялось";
            }
        }

        private static Brush GetStatusBrush(EvacuationRoutesStatus status)
        {
            switch (status)
            {
                case EvacuationRoutesStatus.Ok:
                    return Brushes.ForestGreen;
                case EvacuationRoutesStatus.Warning:
                    return Brushes.Goldenrod;
                case EvacuationRoutesStatus.Error:
                    return Brushes.Firebrick;
                case EvacuationRoutesStatus.Built:
                    return Brushes.DodgerBlue;
                case EvacuationRoutesStatus.PartialBuilt:
                    return Brushes.Goldenrod;
                default:
                    return Brushes.Gray;
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum EvacuationRoutesStatus
    {
        NotChecked = 0,
        Ok = 1,
        Warning = 2,
        Error = 3,
        Built = 4,
        PartialBuilt = 5
    }

    public sealed class EvacuationRoutesStatusUpdate
    {
        public long ElementId { get; set; }
        public EvacuationRoutesStatus Status { get; set; }
        public string StatusText { get; set; }
        public string Message { get; set; }
    }

    public sealed class EvacuationRoutesOperationResult
    {
        public List<EvacuationRoutesStatusUpdate> Updates { get; set; } = new List<EvacuationRoutesStatusUpdate>();
        public List<string> ReportLines { get; set; } = new List<string>();
        public List<EvacuationRoutesProblemGroup> ProblemGroups { get; set; } = new List<EvacuationRoutesProblemGroup>();
    }

    public sealed class EvacuationRoutesProblemGroup
    {
        public long StairElementId { get; set; }
        public List<EvacuationRoutesProblemItem> Items { get; set; } = new List<EvacuationRoutesProblemItem>();
    }

    public sealed class EvacuationRoutesProblemItem
    {
        public string ComponentKind { get; set; }
        public long ComponentElementId { get; set; }
        public long RouteElementId { get; set; }
        public string Message { get; set; }
        public double CurrentLengthMm { get; set; }
        public double CurrentWidthMm { get; set; }
        public double CurrentHeightMm { get; set; }
        public List<EvacuationRoutesProblemTarget> Targets { get; set; } = new List<EvacuationRoutesProblemTarget>();
    }

    public sealed class EvacuationRoutesProblemTarget
    {
        public long ElementId { get; set; }
        public long? LinkInstanceId { get; set; }
        public string DisplayText { get; set; }
    }

    public sealed class EvacuationRoutesCheckRequest
    {
        public long StairElementId { get; set; }
        public long ComponentElementId { get; set; }
        public long RouteElementId { get; set; }
    }

    public sealed class EvacuationRoutesResizeRequest
    {
        public long StairElementId { get; set; }
        public long ComponentElementId { get; set; }
        public long RouteElementId { get; set; }
        public double NewLengthMm { get; set; }
        public double NewWidthMm { get; set; }
        public double NewHeightMm { get; set; }
        public int LengthDirection { get; set; }
        public int WidthDirection { get; set; }
    }

    public sealed class EvacuationRoutesDialogResult
    {
        public int HeightMm { get; }
        public int WidthMm { get; }
        public bool UseRunWidth { get; }
        public bool ConsiderRailings { get; }
        public bool RoundRunWidthDownTo5Mm { get; }
        public bool PickSingleStair { get; }
        public bool AddToEvacuationWorkset { get; }
        public int? EvacuationWorksetId { get; }
        public long? SelectedElementId { get; }
        public List<long> IncludedElementIds { get; }
        public Dictionary<long, bool> UseRunWidthByElementId { get; }

        public EvacuationRoutesDialogResult(
            int heightMm,
            int widthMm,
            bool useRunWidth,
            bool considerRailings,
            bool roundRunWidthDownTo5Mm,
            bool pickSingleStair,
            bool addToEvacuationWorkset,
            int? evacuationWorksetId,
            long? selectedElementId = null,
            IEnumerable<long> includedElementIds = null,
            IDictionary<long, bool> useRunWidthByElementId = null)
        {
            HeightMm = heightMm;
            WidthMm = widthMm;
            UseRunWidth = useRunWidth;
            ConsiderRailings = considerRailings;
            RoundRunWidthDownTo5Mm = roundRunWidthDownTo5Mm;
            PickSingleStair = pickSingleStair;
            AddToEvacuationWorkset = addToEvacuationWorkset;
            EvacuationWorksetId = evacuationWorksetId;
            SelectedElementId = selectedElementId;
            IncludedElementIds = (includedElementIds ?? Enumerable.Empty<long>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();
            UseRunWidthByElementId = (useRunWidthByElementId ?? new Dictionary<long, bool>())
                .Where(x => x.Key > 0)
                .GroupBy(x => x.Key)
                .ToDictionary(x => x.Key, x => x.First().Value);
        }
    }

    public partial class EvacuationRoutesDialog : Window
    {
        public EvacuationRoutesDialogResult Result { get; private set; }

        private const int MinHeightMm = 2100;
        private const int DefaultHeightMm = 2200;
        private static readonly Regex _digitsOnly = new Regex("^[0-9]+$");
        private readonly List<EvacuationRoutesWorksetOption> _evacuationWorksets;
        private readonly ObservableCollection<EvacuationRoutesStairListItem> _stairs;
        private readonly Action<EvacuationRoutesDialogResult> _pickAndBuild;
        private readonly Action<long> _selectElement;
        private readonly Action<EvacuationRoutesDialogResult> _runOperation;
        private readonly Action _pickResizeRoute;
        private List<EvacuationRoutesProblemGroup> _lastProblemGroups = new List<EvacuationRoutesProblemGroup>();
        private EvacuationRoutesReportDialog _reportWindow;
        private bool _restoreReportAfterPick;
        private bool _closingReportWithOwner;
        private long? _pendingSelectedElementId;

        private List<string> _lastReportLines = new List<string>();
        private List<EvacuationRoutesStairListItem> _lastReportStairs = new List<EvacuationRoutesStairListItem>();
        private bool _hasOperationReport;
        private bool _isBusy;

        public EvacuationRoutesDialog(
            IEnumerable<EvacuationRoutesStairListItem> stairs,
            IEnumerable<EvacuationRoutesWorksetOption> evacuationWorksets,
            Action<EvacuationRoutesDialogResult> pickAndBuild,
            Action<long> selectElement,
            Action<EvacuationRoutesDialogResult> runOperation,
            Action pickResizeRoute)
        {
            InitializeComponent();
            Closing += (sender, args) =>
            {
                _closingReportWithOwner = true;
                if (_reportWindow != null)
                    _reportWindow.Close();
            };

            _stairs = new ObservableCollection<EvacuationRoutesStairListItem>((stairs ?? Enumerable.Empty<EvacuationRoutesStairListItem>()).ToList());
            _evacuationWorksets = (evacuationWorksets ?? Enumerable.Empty<EvacuationRoutesWorksetOption>()).ToList();
            _pickAndBuild = pickAndBuild;
            _selectElement = selectElement;
            _runOperation = runOperation;
            _pickResizeRoute = pickResizeRoute;

            TbHeightMm.Text = DefaultHeightMm.ToString(CultureInfo.InvariantCulture);
            TbWidthMm.Text = "1200";
            TbStairCount.Text = $"Найдено лестниц в документе: {_stairs.Count}";

            CmbEvacuationWorksets.ItemsSource = _evacuationWorksets;
            if (_evacuationWorksets.Count > 0)
                CmbEvacuationWorksets.SelectedIndex = 0;

            bool hasEvacuationWorksets = _evacuationWorksets.Count > 0;
            CbAddToEvacuationWorkset.IsChecked = hasEvacuationWorksets;
            CbAddToEvacuationWorkset.IsEnabled = hasEvacuationWorksets;

            CbUseRunWidth.IsChecked = true;
            CbConsiderRailings.IsChecked = true;
            CbRoundRunWidth.IsChecked = true;
            ApplyWidthMode();
            ApplyEvacuationWorksetMode();
            UpdateStatus("Задайте параметры и выберите режим построения.");
        }

        private void PickAndBuild_Click(object sender, RoutedEventArgs e)
        {
            if (!TryCreateResult(pickSingleStair: true, selectedElementId: null, out EvacuationRoutesDialogResult result))
                return;

            ResetStatuses();
            SetBusy("Выберите лестницу в Revit...");
            _pickAndBuild?.Invoke(result);
        }

        private void BuildAll_Click(object sender, RoutedEventArgs e)
        {
            RunOperation(pickSingleStair: false, selectedElementId: null);
        }

        private void PickResizeRoute_Click(object sender, RoutedEventArgs e)
        {
            OpenResizeEditor();
        }

        private void OpenResizeEditor()
        {
            SetBusy("Выберите построенный путь эвакуации в Revit для редактирования...");
            _pickResizeRoute?.Invoke();
        }

        private void SaveReport()
        {
            if (!_hasOperationReport)
            {
                UpdateStatus("Отчёт появится после обработки.");
                return;
            }

            try
            {
                string path = SaveStatusReport();
                if (string.IsNullOrWhiteSpace(path))
                {
                    UpdateStatus("Сохранение отчёта отменено.");
                    return;
                }

                UpdateStatus($"TXT-отчёт сохранён: {path}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Не удалось сохранить TXT-отчёт: {ex.Message}");
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void UseRunWidth_Changed(object sender, RoutedEventArgs e)
        {
            ApplyWidthMode();
        }

        private void EvacuationWorkset_Changed(object sender, RoutedEventArgs e)
        {
            ApplyEvacuationWorksetMode();
        }

        private void RunOperation(bool pickSingleStair, long? selectedElementId)
        {
            if (!TryCreateResult(pickSingleStair, selectedElementId, out EvacuationRoutesDialogResult result))
                return;

            Result = result;
            ResetStatuses();
            SetBusy("Идёт построение...");
            _runOperation?.Invoke(result);
        }

        public void ApplyOperationResult(EvacuationRoutesOperationResult operationResult)
        {
            SetBusy(null, false);

            if (operationResult == null)
                return;

            _lastReportLines = operationResult.ReportLines ?? new List<string>();
            _hasOperationReport = true;
            _lastProblemGroups = operationResult.ProblemGroups ?? new List<EvacuationRoutesProblemGroup>();

            _lastReportStairs = (operationResult.Updates ?? new List<EvacuationRoutesStatusUpdate>())
                .Select(update => update == null ? null : FindItem(update.ElementId))
                .Where(item => item != null)
                .Distinct()
                .ToList();

            foreach (EvacuationRoutesStatusUpdate update in operationResult.Updates ?? new List<EvacuationRoutesStatusUpdate>())
            {
                EvacuationRoutesStairListItem item = FindItem(update.ElementId);
                if (item == null)
                    continue;

                item.SetStatus(update.Status, update.StatusText, update.Message);
            }

            UpdateStatus("Построение завершено. Открыт отчёт по лестницам.");
            ShowReportWindow();
        }

        public void ShowRequestError(string text)
        {
            SetBusy(null, false);
            UpdateStatus(text);
            MessageBox.Show(GetActiveOwnerWindow(), text, "Ошибка",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public void NotifyRequestStatus(string text)
        {
            UpdateStatus(text);
        }

        public void FinishRequest(string text)
        {
            SetBusy(text, false);
        }

        public void SelectRowByElementId(long id)
        {
            EvacuationRoutesStairListItem item = FindItem(id);
            if (item == null)
            {
                UpdateStatus($"Элемент ID {id} не найден в таблице.");
                return;
            }

            _pendingSelectedElementId = item.ElementId;
            _reportWindow?.SelectStair(item);

            UpdateStatus($"Выбрана строка ID {id}.");
        }

        public void RestoreAfterPick()
        {
            if (_restoreReportAfterPick && _reportWindow != null)
            {
                _reportWindow.Show();
                _reportWindow.Activate();
            }
            else
            {
                Show();
                Activate();
            }

            _restoreReportAfterPick = false;
        }

        public void HideForPick()
        {
            _restoreReportAfterPick = _reportWindow != null && _reportWindow.IsVisible;
            if (_restoreReportAfterPick)
                _reportWindow.Hide();
            else
                Hide();
        }

        private bool TryCreateResult(bool pickSingleStair, long? selectedElementId, out EvacuationRoutesDialogResult result)
        {
            result = null;

            if (!TryParsePositiveInt(TbHeightMm.Text, out int heightMm))
            {
                MessageBox.Show(this, "Высота должна быть числом (мм).", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (heightMm < MinHeightMm)
            {
                MessageBox.Show(this, $"Высота не может быть меньше {MinHeightMm} мм.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            bool useRunWidth = CbUseRunWidth.IsChecked == true;
            bool considerRailings = CbConsiderRailings.IsChecked == true;
            bool roundRunWidthDownTo5Mm = CbRoundRunWidth.IsChecked == true;

            int widthMm = 0;
            if (!TryParsePositiveInt(TbWidthMm.Text, out widthMm))
            {
                MessageBox.Show(this, "Ширина должна быть числом (мм).", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!TryGetEvacuationWorksetResult(out bool addToEvacuationWorkset, out int? evacuationWorksetId))
                return false;

            List<long> includedElementIds = _stairs
                .Where(x => x != null)
                .Select(x => x.ElementId)
                .Distinct()
                .ToList();

            if (!pickSingleStair && includedElementIds.Count == 0)
            {
                MessageBox.Show(this, "В документе не найдено лестниц для построения.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            Dictionary<long, bool> useRunWidthByElementId = BuildUseRunWidthByElementId(useRunWidth);
            result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, considerRailings, roundRunWidthDownTo5Mm, pickSingleStair, addToEvacuationWorkset, evacuationWorksetId, selectedElementId, includedElementIds, useRunWidthByElementId);
            return true;
        }

        private void ApplyWidthMode()
        {
            if (string.IsNullOrWhiteSpace(TbWidthMm.Text))
                TbWidthMm.Text = "1200";

            bool manualWidth = (CbUseRunWidth == null || CbUseRunWidth.IsChecked != true)
                && (CbConsiderRailings == null || CbConsiderRailings.IsChecked != true);
            bool enabled = !_isBusy && manualWidth;

            TbWidthMm.IsEnabled = enabled;
            if (LblWidthMm != null)
                LblWidthMm.IsEnabled = enabled;
        }

        private Dictionary<long, bool> BuildUseRunWidthByElementId(bool useRunWidth)
        {
            var result = new Dictionary<long, bool>();
            foreach (EvacuationRoutesStairListItem item in _stairs ?? new ObservableCollection<EvacuationRoutesStairListItem>())
            {
                if (item == null || item.ElementId <= 0)
                    continue;

                result[item.ElementId] = useRunWidth;
            }

            return result;
        }

        private void ApplyEvacuationWorksetMode()
        {
            if (CmbEvacuationWorksets == null)
                return;

            bool canSelect = _evacuationWorksets != null && _evacuationWorksets.Count > 0 && CbAddToEvacuationWorkset.IsChecked == true;
            CmbEvacuationWorksets.IsEnabled = !_isBusy && canSelect;
        }

        private bool TryGetEvacuationWorksetResult(out bool addToEvacuationWorkset, out int? evacuationWorksetId)
        {
            addToEvacuationWorkset = CbAddToEvacuationWorkset.IsChecked == true;
            evacuationWorksetId = null;

            if (!addToEvacuationWorkset)
                return true;

            EvacuationRoutesWorksetOption selected = CmbEvacuationWorksets.SelectedItem as EvacuationRoutesWorksetOption;
            if (selected == null)
            {
                MessageBox.Show(this, "Выберите рабочий набор для путей эвакуации.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            evacuationWorksetId = selected.Id;
            return true;
        }

        private EvacuationRoutesStairListItem FindItem(long id)
        {
            return _stairs.FirstOrDefault(x => x.ElementId == id)
                ?? _stairs.FirstOrDefault(x => x.NestedStairIds != null && x.NestedStairIds.Contains(id));
        }

        private string SaveStatusReport()
        {
            var dialog = new SaveFileDialog
            {
                Title = "Сохранить отчёт",
                FileName = $"KPLN_EvacuationRoutes_Status_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*",
                DefaultExt = ".txt",
                AddExtension = true,
                OverwritePrompt = true
            };

            bool? ok = dialog.ShowDialog(GetActiveOwnerWindow());
            if (ok != true)
                return null;

            string path = dialog.FileName;

            var lines = new List<string>
            {
                "KPLN. Пути эвакуации — отчёт обработки лестниц",
                $"Дата: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                ""
            };

            if (_lastReportLines != null && _lastReportLines.Count > 0)
                lines.AddRange(_lastReportLines);
            else
                lines.Add("Ошибок и пересечений не найдено.");

            File.WriteAllLines(path, lines, Encoding.UTF8);
            return path;
        }

        private void ResetStatuses()
        {
            foreach (EvacuationRoutesStairListItem item in _stairs)
            {
                if (item.HasExistingRoute)
                {
                    EvacuationRoutesStatus status = item.ExistingRouteStatus == EvacuationRoutesStatus.PartialBuilt
                        ? EvacuationRoutesStatus.PartialBuilt
                        : EvacuationRoutesStatus.Built;
                    item.SetStatus(status, null, null);
                }
                else
                {
                    item.SetStatus(EvacuationRoutesStatus.NotChecked, "Не проверялось", null);
                }
            }
        }

        private void SetBusy(string text, bool busy = true)
        {
            _isBusy = busy;
            BtnPickAndBuild.IsEnabled = !busy;
            BtnBuildAll.IsEnabled = !busy;
            BtnPickResizeRoute.IsEnabled = !busy;
            TbHeightMm.IsEnabled = !busy;
            CbUseRunWidth.IsEnabled = !busy;
            CbConsiderRailings.IsEnabled = !busy;
            CbRoundRunWidth.IsEnabled = !busy;
            CbAddToEvacuationWorkset.IsEnabled = !busy && _evacuationWorksets != null && _evacuationWorksets.Count > 0;
            _reportWindow?.SetBusy(busy);
            ApplyWidthMode();
            ApplyEvacuationWorksetMode();

            if (!string.IsNullOrWhiteSpace(text))
                UpdateStatus(text);
        }

        private void ShowReportWindow()
        {
            if (_reportWindow != null)
            {
                _closingReportWithOwner = true;
                _reportWindow.Close();
                _closingReportWithOwner = false;
            }

            EvacuationRoutesReportDialog window = null;
            window = new EvacuationRoutesReportDialog(
                _lastReportStairs,
                elementId =>
                {
                    _selectElement?.Invoke(elementId);
                    UpdateStatus($"Запрошен переход к элементу ID {elementId}.");
                },
                SaveReport,
                OpenResizeEditor);

            window.Closed += (sender, args) =>
            {
                bool closeOwner = ReferenceEquals(_reportWindow, window) && !_closingReportWithOwner;
                if (ReferenceEquals(_reportWindow, window))
                    _reportWindow = null;

                if (closeOwner)
                {
                    _closingReportWithOwner = true;
                    Close();
                }
            };

            _reportWindow = window;
            IntPtr ownerHandle = new WindowInteropHelper(this).Owner;
            if (ownerHandle != IntPtr.Zero)
                new WindowInteropHelper(window) { Owner = ownerHandle };

            if (_pendingSelectedElementId.HasValue)
                window.SelectStair(FindItem(_pendingSelectedElementId.Value));

            Hide();
            window.Show();
            window.Activate();
        }

        public bool TryCreateResizeRequestForRoute(
            long routeElementId,
            long stairElementId,
            long componentElementId,
            double currentLengthMm,
            double currentWidthMm,
            double currentHeightMm,
            out EvacuationRoutesResizeRequest request)
        {
            request = null;

            var item = new EvacuationRoutesProblemItem
            {
                RouteElementId = routeElementId,
                ComponentElementId = componentElementId,
                CurrentLengthMm = currentLengthMm,
                CurrentWidthMm = currentWidthMm,
                CurrentHeightMm = currentHeightMm
            };

            if (!TryShowResizeDialog(item, out double newLengthMm, out double newWidthMm, out double newHeightMm, out int lengthDirection, out int widthDirection))
                return false;

            request = new EvacuationRoutesResizeRequest
            {
                StairElementId = stairElementId,
                ComponentElementId = componentElementId,
                RouteElementId = routeElementId,
                NewLengthMm = newLengthMm,
                NewWidthMm = newWidthMm,
                NewHeightMm = newHeightMm,
                LengthDirection = lengthDirection,
                WidthDirection = widthDirection
            };

            return true;
        }

        private bool TryShowResizeDialog(EvacuationRoutesProblemItem item, out double newLengthMm, out double newWidthMm, out double newHeightMm, out int lengthDirection, out int widthDirection)
        {
            newLengthMm = item?.CurrentLengthMm ?? 0;
            newWidthMm = item?.CurrentWidthMm ?? 0;
            newHeightMm = item?.CurrentHeightMm ?? 0;
            lengthDirection = 0;
            widthDirection = 0;

            var dialog = new EvacuationRoutesResizeDialog(
                item?.RouteElementId ?? 0,
                newLengthMm,
                newWidthMm,
                newHeightMm)
            {
                Owner = GetActiveOwnerWindow()
            };

            if (dialog.ShowDialog() != true)
                return false;

            newLengthMm = dialog.NewLengthMm;
            newWidthMm = dialog.NewWidthMm;
            newHeightMm = dialog.NewHeightMm;
            lengthDirection = dialog.LengthDirection;
            widthDirection = dialog.WidthDirection;
            return true;
        }

        public void ShowRouteCheckResult(string text)
        {
            MessageBox.Show(GetActiveOwnerWindow(), string.IsNullOrWhiteSpace(text) ? "Пересечений не найдено." : text,
                "KPLN. Проверка пересечений", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public void MarkStairFixed(long stairId)
        {
            EvacuationRoutesStairListItem item = FindItem(stairId);
            if (item == null)
                return;

            item.HasExistingRoute = true;
            item.ExistingRouteStatus = EvacuationRoutesStatus.Built;
            item.SetStatus(EvacuationRoutesStatus.Ok, "ОК (Исправлено)", null);
        }

        public void UpdateRouteDimensions(long routeElementId, double lengthMm, double widthMm, double heightMm)
        {
            if (routeElementId <= 0 || _lastProblemGroups == null)
                return;

            foreach (EvacuationRoutesProblemItem item in _lastProblemGroups
                .Where(x => x != null && x.Items != null)
                .SelectMany(x => x.Items)
                .Where(x => x != null && x.RouteElementId == routeElementId))
            {
                if (lengthMm > 0) item.CurrentLengthMm = lengthMm;
                if (widthMm > 0) item.CurrentWidthMm = widthMm;
                if (heightMm > 0) item.CurrentHeightMm = heightMm;
            }
        }

        private void UpdateStatus(string text)
        {
            if (TbStatus != null)
                TbStatus.Text = text ?? "";
            _reportWindow?.SetStatus(text);
        }

        private Window GetActiveOwnerWindow()
        {
            if (_reportWindow != null && _reportWindow.IsVisible)
                return _reportWindow;

            return this;
        }

        private void DigitsOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !_digitsOnly.IsMatch(e.Text);
        }

        private void DigitsOnly_OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(DataFormats.Text))
            {
                e.CancelCommand();
                return;
            }

            var text = e.DataObject.GetData(DataFormats.Text) as string ?? "";
            if (!_digitsOnly.IsMatch(text))
                e.CancelCommand();
        }

        private static bool TryParsePositiveInt(string s, out int value)
        {
            return int.TryParse((s ?? "").Trim(), out value) && value > 0;
        }

    }
}