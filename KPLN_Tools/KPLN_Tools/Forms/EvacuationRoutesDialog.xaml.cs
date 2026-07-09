using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
        Error = 3
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

        private List<string> _lastReportLines = new List<string>();
        private bool _hasOperationReport;
        private bool _suppressSelectionAction;
        private bool _isBusy;

        public EvacuationRoutesDialog(
            IEnumerable<EvacuationRoutesStairListItem> stairs,
            IEnumerable<EvacuationRoutesWorksetOption> evacuationWorksets,
            Action<EvacuationRoutesDialogResult> pickAndBuild,
            Action<long> selectElement,
            Action<EvacuationRoutesDialogResult> runOperation)
        {
            InitializeComponent();

            _stairs = new ObservableCollection<EvacuationRoutesStairListItem>((stairs ?? Enumerable.Empty<EvacuationRoutesStairListItem>()).ToList());
            _evacuationWorksets = (evacuationWorksets ?? Enumerable.Empty<EvacuationRoutesWorksetOption>()).ToList();
            _pickAndBuild = pickAndBuild;
            _selectElement = selectElement;
            _runOperation = runOperation;

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
            UpdateStatus("Выберите строку или запустите обработку.");
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
                item.SetStatus(EvacuationRoutesStatus.NotChecked, "Не проверялось", null);
        }

        private void SetBusy(string text, bool busy = true)
        {
            _isBusy = busy;
            BtnSaveReport.IsEnabled = !busy;
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
    }
}