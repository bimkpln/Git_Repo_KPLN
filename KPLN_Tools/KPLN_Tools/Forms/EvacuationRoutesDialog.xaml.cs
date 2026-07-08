using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Collections.Generic;
using System.Linq;

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

    public sealed class EvacuationRoutesDialogResult
    {
        public int HeightMm { get; }
        public int WidthMm { get; }
        public bool UseRunWidth { get; }
        public bool PickSingleStair { get; }
        public bool AddToEvacuationWorkset { get; }
        public int? EvacuationWorksetId { get; }

        public EvacuationRoutesDialogResult(int heightMm, int widthMm, bool useRunWidth, bool pickSingleStair, bool addToEvacuationWorkset, int? evacuationWorksetId)
        {
            HeightMm = heightMm;
            WidthMm = widthMm;
            UseRunWidth = useRunWidth;
            PickSingleStair = pickSingleStair;
            AddToEvacuationWorkset = addToEvacuationWorkset;
            EvacuationWorksetId = evacuationWorksetId;
        }
    }

    public partial class EvacuationRoutesDialog : Window
    {
        public EvacuationRoutesDialogResult Result { get; private set; }

        private const int MinHeightMm = 2100;
        private static readonly Regex _digitsOnly = new Regex("^[0-9]+$");
        private readonly List<EvacuationRoutesWorksetOption> _evacuationWorksets;

        public EvacuationRoutesDialog(int stairsCount, IEnumerable<EvacuationRoutesWorksetOption> evacuationWorksets)
        {
            InitializeComponent();

            _evacuationWorksets = (evacuationWorksets ?? Enumerable.Empty<EvacuationRoutesWorksetOption>()).ToList();

            TbStairsInfo.Text = $"Найдено лестниц в документе: {stairsCount}";
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
        }

        private void PickOne_Click(object sender, RoutedEventArgs e)
        {
            if (!TryParsePositiveInt(TbHeightMm.Text, out int heightMm))
            {
                MessageBox.Show(this, "Высота должна быть числом (мм).", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (heightMm < MinHeightMm)
            {
                MessageBox.Show(this, $"Высота не может быть меньше {MinHeightMm} мм.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool useRunWidth = CbUseRunWidth.IsChecked == true;

            int widthMm = 0;
            if (!useRunWidth)
            {
                if (!TryParsePositiveInt(TbWidthMm.Text, out widthMm))
                {
                    MessageBox.Show(this, "Ширина должна быть числом (мм).", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            if (!TryGetEvacuationWorksetResult(out bool addToEvacuationWorkset, out int? evacuationWorksetId))
                return;

            Result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, pickSingleStair: true, addToEvacuationWorkset, evacuationWorksetId);
            DialogResult = true;
            Close();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!TryParsePositiveInt(TbHeightMm.Text, out int heightMm))
            {
                MessageBox.Show(this, "Высота должна быть числом (мм).", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (heightMm < MinHeightMm)
            {
                MessageBox.Show(this, $"Высота не может быть меньше {MinHeightMm} мм.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            bool useRunWidth = CbUseRunWidth.IsChecked == true;

            int widthMm = 0;
            if (!useRunWidth)
            {
                if (!TryParsePositiveInt(TbWidthMm.Text, out widthMm))
                {
                    MessageBox.Show(this, "Ширина должна быть числом (мм).", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            if (!TryGetEvacuationWorksetResult(out bool addToEvacuationWorkset, out int? evacuationWorksetId))
                return;

            Result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, pickSingleStair: false, addToEvacuationWorkset, evacuationWorksetId);
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
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

        private void ApplyWidthMode()
        {
            bool useRunWidth = CbUseRunWidth.IsChecked == true;
            TbWidthMm.IsEnabled = !useRunWidth;
            if (useRunWidth)
                TbWidthMm.Text = "";
        }

        private void ApplyEvacuationWorksetMode()
        {
            if (CmbEvacuationWorksets == null)
                return;

            bool canSelect = _evacuationWorksets != null && _evacuationWorksets.Count > 0 && CbAddToEvacuationWorkset.IsChecked == true;
            CmbEvacuationWorksets.IsEnabled = canSelect;
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