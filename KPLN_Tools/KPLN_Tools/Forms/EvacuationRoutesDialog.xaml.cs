using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.ExternalCommands.UI
{
    public sealed class EvacuationRoutesDialogResult
    {
        public int HeightMm { get; }
        public int WidthMm { get; }
        public bool UseRunWidth { get; }
        public bool PickSingleStair { get; }  

        public EvacuationRoutesDialogResult(int heightMm, int widthMm, bool useRunWidth, bool pickSingleStair)
        {
            HeightMm = heightMm;
            WidthMm = widthMm;
            UseRunWidth = useRunWidth;
            PickSingleStair = pickSingleStair;
        }
    }

    public partial class EvacuationRoutesDialog : Window
    {
        public EvacuationRoutesDialogResult Result { get; private set; }

        private const int MinHeightMm = 2100;
        private static readonly Regex _digitsOnly = new Regex("^[0-9]+$");

        public EvacuationRoutesDialog(int stairsCount)
        {
            InitializeComponent();

            TbStairsInfo.Text = $"Найдено лестниц в документе: {stairsCount}";
            TbHeightMm.Text = MinHeightMm.ToString();
            TbWidthMm.Text = "1200";

            CbUseRunWidth.IsChecked = true;
            ApplyWidthMode();
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

            Result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, pickSingleStair: true);
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

            Result = new EvacuationRoutesDialogResult(heightMm, widthMm, useRunWidth, pickSingleStair: false);
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

        private void ApplyWidthMode()
        {
            bool useRunWidth = CbUseRunWidth.IsChecked == true;
            TbWidthMm.IsEnabled = !useRunWidth;
            if (useRunWidth)
                TbWidthMm.Text = "";
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