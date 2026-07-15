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

    public sealed class EvacuationRoutesTrimRequest
    {
        public long StairElementId { get; set; }
        public long ComponentElementId { get; set; }
        public long RouteElementId { get; set; }
        public List<long> IntersectingElementIds { get; set; } = new List<long>();
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
        public bool PickSingleStair { get; }
        public bool AddToEvacuationWorkset { get; }
        public int? EvacuationWorksetId { get; }
        public long? SelectedElementId { get; }
        public List<long> IncludedElementIds { get; }

        public EvacuationRoutesDialogResult(int heightMm, int widthMm, bool useRunWidth, bool pickSingleStair, bool addToEvacuationWorkset, int? evacuationWorksetId, long? selectedElementId = null, IEnumerable<long> includedElementIds = null)
        {
            HeightMm = heightMm;
            WidthMm = widthMm;
            UseRunWidth = useRunWidth;
            PickSingleStair = pickSingleStair;
            AddToEvacuationWorkset = addToEvacuationWorkset;
            EvacuationWorksetId = evacuationWorksetId;
            SelectedElementId = selectedElementId;
            IncludedElementIds = (includedElementIds ?? Enumerable.Empty<long>())
                .Where(x => x > 0)
                .Distinct()
                .ToList();
        }
    }

    public partial class EvacuationRoutesDialog : Window
    {
        public EvacuationRoutesDialogResult Result { get; private set; }

        private const int MinHeightMm = 2100;
        private static readonly Regex _digitsOnly = new Regex("^[0-9]+$");
        private readonly List<EvacuationRoutesWorksetOption> _evacuationWorksets;
        private readonly ObservableCollection<EvacuationRoutesStairListItem> _stairs;
        private readonly Action<EvacuationRoutesDialogResult> _pickAndBuild;
        private readonly Action<long> _selectElement;
        private readonly Action<EvacuationRoutesDialogResult> _runOperation;
        private readonly Action<EvacuationRoutesTrimRequest> _trimRoute;
        private readonly Action<EvacuationRoutesResizeRequest> _resizeRoute;
        private List<EvacuationRoutesProblemGroup> _lastProblemGroups = new List<EvacuationRoutesProblemGroup>();

        private List<string> _lastReportLines = new List<string>();
        private bool _hasOperationReport;
        private bool _suppressSelectionAction;
        private bool _isBusy;

        public EvacuationRoutesDialog(
            IEnumerable<EvacuationRoutesStairListItem> stairs,
            IEnumerable<EvacuationRoutesWorksetOption> evacuationWorksets,
            Action<EvacuationRoutesDialogResult> pickAndBuild,
            Action<long> selectElement,
            Action<EvacuationRoutesDialogResult> runOperation,
            Action<EvacuationRoutesTrimRequest> trimRoute,
            Action<EvacuationRoutesResizeRequest> resizeRoute)
        {
            InitializeComponent();

            _stairs = new ObservableCollection<EvacuationRoutesStairListItem>((stairs ?? Enumerable.Empty<EvacuationRoutesStairListItem>()).ToList());
            _evacuationWorksets = (evacuationWorksets ?? Enumerable.Empty<EvacuationRoutesWorksetOption>()).ToList();
            _pickAndBuild = pickAndBuild;
            _selectElement = selectElement;
            _runOperation = runOperation;
            _trimRoute = trimRoute;
            _resizeRoute = resizeRoute;

            foreach (EvacuationRoutesStairListItem item in _stairs)
            {
                if (item != null)
                    item.PropertyChanged += StairItem_PropertyChanged;
            }

            DgStairs.ItemsSource = _stairs;

            TbHeightMm.Text = MinHeightMm.ToString();
            TbWidthMm.Text = "1200";

            CmbEvacuationWorksets.ItemsSource = _evacuationWorksets;
            if (_evacuationWorksets.Count > 0)
                CmbEvacuationWorksets.SelectedIndex = 0;

            bool hasEvacuationWorksets = _evacuationWorksets.Count > 0;
            CbAddToEvacuationWorkset.IsChecked = hasEvacuationWorksets;
            CbAddToEvacuationWorkset.IsEnabled = hasEvacuationWorksets;

            CbUseRunWidth.IsChecked = true;
            ApplyWidthMode();
            ApplyEvacuationWorksetMode();
            UpdateToggleSelectionButton();
            UpdateStatus("Выберите строку или запустите обработку.");
        }

        private void StairItem_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (string.Equals(e.PropertyName, nameof(EvacuationRoutesStairListItem.IsIncluded), StringComparison.Ordinal))
                UpdateToggleSelectionButton();
        }

        private void DgStairs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_suppressSelectionAction)
                return;

            EvacuationRoutesStairListItem item = GetSelectedItem();
            if (item == null)
                return;

            _selectElement?.Invoke(item.ElementId);
            UpdateStatus($"Запрошен переход к элементу ID {item.ElementId}.");
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

        private void ToggleSelection_Click(object sender, RoutedEventArgs e)
        {
            bool allIncluded = _stairs.Count > 0 && _stairs.All(x => x != null && x.IsIncluded);
            SetAllIncluded(!allIncluded);
        }

        private void SaveReport_Click(object sender, RoutedEventArgs e)
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
            BtnSaveReport.Visibility = Visibility.Visible;

            foreach (EvacuationRoutesStatusUpdate update in operationResult.Updates ?? new List<EvacuationRoutesStatusUpdate>())
            {
                EvacuationRoutesStairListItem item = FindItem(update.ElementId);
                if (item == null)
                    continue;

                item.SetStatus(update.Status, update.StatusText, update.Message);
            }

            UpdateStatus("Построение завершено. Подробности в статусах строк.");
            ShowOperationSummary(operationResult);
        }

        public void ShowRequestError(string text)
        {
            SetBusy(null, false);
            UpdateStatus(text);
            MessageBox.Show(this, text, "Ошибка",
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

            _suppressSelectionAction = true;
            try
            {
                DgStairs.SelectedItem = item;
                DgStairs.ScrollIntoView(item);
            }
            finally
            {
                _suppressSelectionAction = false;
            }

            UpdateStatus($"Выбрана строка ID {id}.");
        }

        public void RestoreAfterPick()
        {
            Show();
            Activate();
        }

        public void HideForPick()
        {
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

            int widthMm = 0;
            if (!useRunWidth)
            {
                if (!TryParsePositiveInt(TbWidthMm.Text, out widthMm))
                {
                    MessageBox.Show(this, "Ширина должна быть числом (мм).", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
            }

            if (!TryGetEvacuationWorksetResult(out bool addToEvacuationWorkset, out int? evacuationWorksetId))
                return false;

            List<long> includedElementIds = _stairs
                .Where(x => x != null && x.IsIncluded)
                .Select(x => x.ElementId)
                .Distinct()
                .ToList();

            if (!pickSingleStair && includedElementIds.Count == 0)
            {
                MessageBox.Show(this, "Отметьте хотя бы одну лестницу для построения.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, pickSingleStair, addToEvacuationWorkset, evacuationWorksetId, selectedElementId, includedElementIds);
            return true;
        }

        private void ApplyWidthMode()
        {
            bool useRunWidth = CbUseRunWidth.IsChecked == true;
            TbWidthMm.IsEnabled = !_isBusy && !useRunWidth;
            if (useRunWidth)
                TbWidthMm.Text = "";
            else if (string.IsNullOrWhiteSpace(TbWidthMm.Text))
                TbWidthMm.Text = "1200";
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

        private EvacuationRoutesStairListItem GetSelectedItem()
        {
            return DgStairs.SelectedItem as EvacuationRoutesStairListItem;
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

            bool? ok = dialog.ShowDialog(this);
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
            BtnSaveReport.IsEnabled = !busy;
            BtnToggleSelection.IsEnabled = !busy;
            BtnPickAndBuild.IsEnabled = !busy;
            BtnBuildAll.IsEnabled = !busy;
            DgStairs.IsEnabled = !busy;
            TbHeightMm.IsEnabled = !busy;
            CbUseRunWidth.IsEnabled = !busy;
            CbAddToEvacuationWorkset.IsEnabled = !busy && _evacuationWorksets != null && _evacuationWorksets.Count > 0;
            ApplyWidthMode();
            ApplyEvacuationWorksetMode();

            if (!string.IsNullOrWhiteSpace(text))
                UpdateStatus(text);
        }

        private void SetAllIncluded(bool included)
        {
            foreach (EvacuationRoutesStairListItem item in _stairs)
            {
                if (item != null)
                    item.IsIncluded = included;
            }

            UpdateStatus(included ? "Все лестницы отмечены для построения." : "Выделение со всех лестниц снято.");
            UpdateToggleSelectionButton();
        }

        private void UpdateToggleSelectionButton()
        {
            if (BtnToggleSelection == null)
                return;

            bool allIncluded = _stairs != null && _stairs.Count > 0 && _stairs.All(x => x != null && x.IsIncluded);
            BtnToggleSelection.Content = allIncluded ? "Снять выделение" : "Выделить всё";
        }

        private void ShowOperationSummary(EvacuationRoutesOperationResult operationResult)
        {
            List<string> lines = operationResult == null
                ? new List<string>()
                : operationResult.ReportLines ?? new List<string>();

            List<EvacuationRoutesProblemGroup> groups = operationResult == null
                ? new List<EvacuationRoutesProblemGroup>()
                : operationResult.ProblemGroups ?? new List<EvacuationRoutesProblemGroup>();
            _lastProblemGroups = groups;

            string body = lines.Count == 0
                ? "Ошибок и пересечений не найдено."
                : string.Join(Environment.NewLine, lines);

            FrameworkElement content = groups.Count > 0
                ? BuildProblemGroupsContent(groups)
                : BuildPlainProblemText(body);

            var closeButton = new Button
            {
                Content = "Закрыть",
                Width = 90,
                Height = 28,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(12, 0, 12, 12)
            };

            var footer = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(244, 246, 248)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 228, 232)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(0, 10, 0, 0),
                Child = closeButton
            };

            var grid = new Grid
            {
                Background = new SolidColorBrush(Color.FromRgb(244, 246, 248))
            };
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(content, 0);
            Grid.SetRow(footer, 1);
            grid.Children.Add(content);
            grid.Children.Add(footer);

            var window = new Window
            {
                Title = "KPLN. Проблемы построения путей эвакуации",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Width = 940,
                Height = 600,
                MinWidth = 760,
                MinHeight = 420,
                ResizeMode = ResizeMode.CanResize,
                Content = grid
            };

            closeButton.Click += (sender, args) => window.Close();
            window.Show();
        }

        private static FrameworkElement BuildPlainProblemText(string body)
        {
            return new TextBox
            {
                Text = body,
                IsReadOnly = true,
                TextWrapping = TextWrapping.Wrap,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Margin = new Thickness(12),
                FontFamily = new FontFamily("Consolas")
            };
        }

        private FrameworkElement BuildProblemGroupsContent(List<EvacuationRoutesProblemGroup> groups)
        {
            var stack = new StackPanel { Margin = new Thickness(12) };

            foreach (EvacuationRoutesProblemGroup group in groups ?? new List<EvacuationRoutesProblemGroup>())
            {
                if (group == null)
                    continue;

                stack.Children.Add(CreateProblemGroupPanel(group));
            }

            if (stack.Children.Count == 0)
                stack.Children.Add(new TextBlock
                {
                    Text = "Ошибок и пересечений не найдено.",
                    TextWrapping = TextWrapping.Wrap,
                    Margin = new Thickness(8)
                });

            return new ScrollViewer
            {
                Content = stack,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
        }

        private FrameworkElement CreateProblemGroupPanel(EvacuationRoutesProblemGroup group)
        {
            var panel = new StackPanel();

            var headerGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            headerGrid.Children.Add(new TextBlock
            {
                Text = $"ЛЕСТНИЦА ID {group.StairElementId}",
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                FontSize = 14,
                Foreground = new SolidColorBrush(Color.FromRgb(35, 45, 58))
            });

            var headerActions = new StackPanel { Orientation = Orientation.Horizontal };
            headerActions.Children.Add(CreateSelectButton("Лестница", group.StairElementId));

            long firstRouteId = (group.Items ?? new List<EvacuationRoutesProblemItem>())
                .Where(x => x != null && x.RouteElementId > 0)
                .Select(x => x.RouteElementId)
                .FirstOrDefault();

            if (firstRouteId > 0)
                headerActions.Children.Add(CreateSelectButton("Путь", firstRouteId));

            Grid.SetColumn(headerActions, 1);
            headerGrid.Children.Add(headerActions);
            panel.Children.Add(headerGrid);

            foreach (EvacuationRoutesProblemItem item in group.Items ?? new List<EvacuationRoutesProblemItem>())
            {
                if (item != null)
                    panel.Children.Add(CreateProblemItemPanel(group, item));
            }

            return new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 224, 231)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 12),
                Child = panel
            };
        }

        private FrameworkElement CreateProblemItemPanel(EvacuationRoutesProblemGroup group, EvacuationRoutesProblemItem item)
        {
            var panel = new StackPanel();
            var row = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            string label = string.IsNullOrWhiteSpace(item.ComponentKind) ? "Элемент" : item.ComponentKind;
            if (item.ComponentElementId > 0)
                label += $" ID {item.ComponentElementId}";

            row.Children.Add(new TextBlock
            {
                Text = $"{label} - {item.Message}",
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = new SolidColorBrush(Color.FromRgb(28, 36, 48)),
                Margin = new Thickness(0, 0, 12, 0)
            });

            var actions = new StackPanel { Orientation = Orientation.Horizontal };

            if (item.ComponentElementId > 0)
                actions.Children.Add(CreateSelectButton("Элемент", item.ComponentElementId));

            if (item.RouteElementId > 0)
                actions.Children.Add(CreateSelectButton("Путь", item.RouteElementId));

            Button trimButton = new Button
            {
                Content = "✂",
                ToolTip = "Обрезать",
                Height = 24,
                Width = 28,
                Margin = new Thickness(4, 0, 0, 0),
                IsEnabled = item.RouteElementId > 0 && item.Targets != null && item.Targets.Any(x => x != null && x.ElementId > 0 && !x.LinkInstanceId.HasValue)
            };
            trimButton.Click += (sender, args) => TrimProblemItem(group, item);
            actions.Children.Add(trimButton);

            Button resizeButton = new Button
            {
                Content = "↔",
                ToolTip = "Изменить габариты",
                Height = 24,
                Width = 28,
                Margin = new Thickness(4, 0, 0, 0),
                IsEnabled = item.RouteElementId > 0
            };
            resizeButton.Click += (sender, args) => ResizeProblemItem(group, item);
            actions.Children.Add(resizeButton);

            Grid.SetColumn(actions, 1);
            row.Children.Add(actions);

            panel.Children.Add(row);

            foreach (EvacuationRoutesProblemTarget target in item.Targets ?? new List<EvacuationRoutesProblemTarget>())
            {
                if (target == null)
                    continue;

                var targetRow = new Grid { Margin = new Thickness(14, 4, 0, 0) };
                targetRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                targetRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                targetRow.Children.Add(new TextBlock
                {
                    Text = target.DisplayText,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalAlignment = VerticalAlignment.Center,
                    Foreground = new SolidColorBrush(Color.FromRgb(72, 84, 98)),
                    Margin = new Thickness(0, 0, 10, 0)
                });

                long selectableId = target.LinkInstanceId.HasValue ? target.LinkInstanceId.Value : target.ElementId;
                if (selectableId > 0)
                {
                    Button targetButton = CreateSelectButton("Выбрать", selectableId);
                    Grid.SetColumn(targetButton, 1);
                    targetRow.Children.Add(targetButton);
                }

                panel.Children.Add(targetRow);
            }

            return new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(230, 235, 241)),
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb(250, 251, 252)),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10),
                Margin = new Thickness(0, 0, 0, 8),
                Child = panel
            };
        }

        private Button CreateSelectButton(string text, long elementId)
        {
            var button = new Button
            {
                Content = text,
                Height = 24,
                MinWidth = 70,
                Margin = new Thickness(4, 0, 0, 0),
                IsEnabled = elementId > 0
            };

            button.Click += (sender, args) => SelectProblemElement(elementId);
            return button;
        }

        private void SelectProblemElement(long elementId)
        {
            if (elementId <= 0)
                return;

            _selectElement?.Invoke(elementId);
            UpdateStatus($"Запрошен переход к элементу ID {elementId}.");
        }

        private void TrimProblemItem(EvacuationRoutesProblemGroup group, EvacuationRoutesProblemItem item)
        {
            if (_trimRoute == null || item == null || item.RouteElementId <= 0)
                return;

            var request = new EvacuationRoutesTrimRequest
            {
                StairElementId = group == null ? 0 : group.StairElementId,
                ComponentElementId = item.ComponentElementId,
                RouteElementId = item.RouteElementId,
                IntersectingElementIds = (item.Targets ?? new List<EvacuationRoutesProblemTarget>())
                    .Where(x => x != null && x.ElementId > 0 && !x.LinkInstanceId.HasValue)
                    .Select(x => x.ElementId)
                    .Distinct()
                    .ToList()
            };

            if (request.IntersectingElementIds.Count == 0)
            {
                UpdateStatus("Для обрезки нет пересекающих элементов в основном файле.");
                ShowRouteCheckResult(null);
                return;
            }

            _trimRoute(request);
            UpdateStatus($"Запрошена обрезка пути ID {item.RouteElementId}.");
        }

        private void ResizeProblemItem(EvacuationRoutesProblemGroup group, EvacuationRoutesProblemItem item)
        {
            if (_resizeRoute == null || item == null || item.RouteElementId <= 0)
                return;

            if (!TryShowResizeDialog(item, out double newLengthMm, out double newWidthMm, out double newHeightMm, out int lengthDirection, out int widthDirection))
                return;

            _resizeRoute(new EvacuationRoutesResizeRequest
            {
                StairElementId = group == null ? 0 : group.StairElementId,
                ComponentElementId = item.ComponentElementId,
                RouteElementId = item.RouteElementId,
                NewLengthMm = newLengthMm,
                NewWidthMm = newWidthMm,
                NewHeightMm = newHeightMm,
                LengthDirection = lengthDirection,
                WidthDirection = widthDirection
            });

            UpdateStatus($"Запрошено изменение габаритов пути ID {item.RouteElementId}.");
        }

        private bool TryShowResizeDialog(EvacuationRoutesProblemItem item, out double newLengthMm, out double newWidthMm, out double newHeightMm, out int lengthDirection, out int widthDirection)
        {
            newLengthMm = item?.CurrentLengthMm ?? 0;
            newWidthMm = item?.CurrentWidthMm ?? 0;
            newHeightMm = item?.CurrentHeightMm ?? 0;
            lengthDirection = 0;
            widthDirection = 0;

            var lengthBox = CreateDimensionTextBox(newLengthMm);
            var widthBox = CreateDimensionTextBox(newWidthMm);
            var heightBox = CreateDimensionTextBox(newHeightMm);
            var lengthDirectionBox = CreateDirectionComboBox(false);
            var widthDirectionBox = CreateDirectionComboBox(true);

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(190) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            for (int i = 0; i < 4; i++)
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var dimensionHeader = new TextBlock
            {
                Text = "Габариты",
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(45, 55, 72)),
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(dimensionHeader, 0);
            Grid.SetColumnSpan(dimensionHeader, 2);
            grid.Children.Add(dimensionHeader);

            var directionHeader = new TextBlock
            {
                Text = "Изменять",
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(45, 55, 72)),
                Margin = new Thickness(0, 0, 0, 8)
            };
            Grid.SetRow(directionHeader, 0);
            Grid.SetColumn(directionHeader, 3);
            grid.Children.Add(directionHeader);

            AddDimensionRow(grid, 1, "Длина, мм", lengthBox, lengthDirectionBox);
            AddDimensionRow(grid, 2, "Ширина, мм", widthBox, widthDirectionBox);
            AddDimensionRow(grid, 3, "Высота, мм", heightBox);

            var note = new TextBlock
            {
                Text = "Меньшая/большая грань считается по локальной оси пути." + Environment.NewLine + "При выборе 'обе' центр остаётся на месте.",
                Foreground = new SolidColorBrush(Color.FromRgb(70, 84, 103)),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0)
            };

            var noteBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(239, 246, 255)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(191, 219, 254)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(10, 8, 10, 8),
                Margin = new Thickness(0, 12, 0, 0),
                Child = note
            };

            var body = new StackPanel();
            body.Children.Add(grid);
            body.Children.Add(noteBorder);

            var bodyCard = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 229, 236)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(14),
                Margin = new Thickness(14),
                Child = body
            };

            var okButton = new Button { Content = "Изменить", Width = 96, Height = 30, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "Отмена", Width = 86, Height = 30, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(14, 10, 14, 10) };
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);

            var footer = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 229, 236)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Child = buttons
            };

            var header = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(224, 229, 236)),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Padding = new Thickness(14, 12, 14, 10),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = $"Путь эвакуации ID {item.RouteElementId}",
                            FontWeight = FontWeights.Bold,
                            FontSize = 13,
                            Foreground = new SolidColorBrush(Color.FromRgb(24, 32, 44))
                        },
                        new TextBlock
                        {
                            Text = "Изменение габаритов без перестроения формы",
                            Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)),
                            Margin = new Thickness(0, 3, 0, 0)
                        }
                    }
                }
            };

            var root = new Grid { Background = new SolidColorBrush(Color.FromRgb(246, 248, 251)) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Grid.SetRow(header, 0);
            Grid.SetRow(bodyCard, 1);
            Grid.SetRow(footer, 2);
            root.Children.Add(header);
            root.Children.Add(bodyCard);
            root.Children.Add(footer);

            var window = new Window
            {
                Title = "KPLN. Изменить габариты пути эвакуации",
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Width = 610,
                Height = 390,
                Content = root
            };

            bool accepted = false;
            double acceptedLengthMm = newLengthMm;
            double acceptedWidthMm = newWidthMm;
            double acceptedHeightMm = newHeightMm;
            int acceptedLengthDirection = lengthDirection;
            int acceptedWidthDirection = widthDirection;
            okButton.Click += (sender, args) =>
            {
                if (!TryParsePositiveDouble(lengthBox.Text, out double length) ||
                    !TryParsePositiveDouble(widthBox.Text, out double width) ||
                    !TryParsePositiveDouble(heightBox.Text, out double height))
                {
                    MessageBox.Show(window, "Все габариты должны быть положительными числами в мм.", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                acceptedLengthMm = length;
                acceptedWidthMm = width;
                acceptedHeightMm = height;
                acceptedLengthDirection = GetSelectedDirection(lengthDirectionBox);
                acceptedWidthDirection = GetSelectedDirection(widthDirectionBox);
                accepted = true;
                window.Close();
            };

            window.ShowDialog();
            if (accepted)
            {
                newLengthMm = acceptedLengthMm;
                newWidthMm = acceptedWidthMm;
                newHeightMm = acceptedHeightMm;
                lengthDirection = acceptedLengthDirection;
                widthDirection = acceptedWidthDirection;
            }

            return accepted;
        }

        private static ComboBox CreateDirectionComboBox(bool isWidth)
        {
            var comboBox = new ComboBox
            {
                Height = 26,
                Margin = new Thickness(0, 0, 0, 6),
                VerticalContentAlignment = VerticalAlignment.Center,
                SelectedValuePath = "Tag"
            };

            comboBox.Items.Add(new ComboBoxItem { Content = isWidth ? "Обе боковые грани" : "Обе торцевые грани", Tag = 0 });
            comboBox.Items.Add(new ComboBoxItem { Content = isWidth ? "Меньшая боковая грань" : "Меньшая торцевая грань", Tag = -1 });
            comboBox.Items.Add(new ComboBoxItem { Content = isWidth ? "Большая боковая грань" : "Большая торцевая грань", Tag = 1 });
            comboBox.SelectedIndex = 0;
            return comboBox;
        }

        private static TextBox CreateDimensionTextBox(double valueMm)
        {
            return new TextBox
            {
                Text = valueMm > 0 ? Math.Round(valueMm, 0).ToString(CultureInfo.InvariantCulture) : "",
                Height = 26,
                Margin = new Thickness(0, 0, 8, 6),
                VerticalContentAlignment = VerticalAlignment.Center
            };
        }

        private static void AddDimensionRow(Grid grid, int row, string label, TextBox textBox, ComboBox directionBox = null)
        {
            var text = new TextBlock
            {
                Text = label,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 6)
            };

            Grid.SetRow(text, row);
            Grid.SetColumn(text, 0);
            Grid.SetRow(textBox, row);
            Grid.SetColumn(textBox, 1);
            grid.Children.Add(text);
            grid.Children.Add(textBox);

            if (directionBox != null)
            {
                Grid.SetRow(directionBox, row);
                Grid.SetColumn(directionBox, 3);
                grid.Children.Add(directionBox);
            }
        }

        public void ShowRouteCheckResult(string text)
        {
            MessageBox.Show(this, string.IsNullOrWhiteSpace(text) ? "Пересечений не найдено." : text,
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

        private static bool TryParsePositiveDouble(string s, out double value)
        {
            s = (s ?? "").Trim().Replace(',', '.');
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && value > 0.0;
        }

        private static int GetSelectedDirection(ComboBox comboBox)
        {
            ComboBoxItem item = comboBox == null ? null : comboBox.SelectedItem as ComboBoxItem;
            if (item == null)
                return 0;

            if (item.Tag is int value)
                return value;

            return 0;
        }
    }
}