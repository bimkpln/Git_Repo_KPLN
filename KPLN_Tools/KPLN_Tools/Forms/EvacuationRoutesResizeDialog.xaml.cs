using System;
using System.Globalization;
using System.Windows;

namespace KPLN_Tools.ExternalCommands.UI
{
    public partial class EvacuationRoutesResizeDialog : Window
    {
        public double NewLengthMm { get; private set; }
        public double NewWidthMm { get; private set; }
        public double NewHeightMm { get; private set; }
        public int LengthDirection { get; private set; }
        public int WidthDirection { get; private set; }

        public EvacuationRoutesResizeDialog(
            long routeElementId,
            double currentLengthMm,
            double currentWidthMm,
            double currentHeightMm)
        {
            InitializeComponent();

            TbRouteTitle.Text = routeElementId > 0
                ? $"Путь эвакуации ID {routeElementId}"
                : "Путь эвакуации";
            TbLengthMm.Text = FormatDimension(currentLengthMm);
            TbWidthMm.Text = FormatDimension(currentWidthMm);
            TbHeightMm.Text = FormatDimension(currentHeightMm);
        }

        private void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (!TryParsePositiveDouble(TbLengthMm.Text, out double length) ||
                !TryParsePositiveDouble(TbWidthMm.Text, out double width) ||
                !TryParsePositiveDouble(TbHeightMm.Text, out double height))
            {
                MessageBox.Show(this,
                    "Все габариты должны быть положительными числами в мм.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            NewLengthMm = length;
            NewWidthMm = width;
            NewHeightMm = height;
            LengthDirection = GetDirection(CmbLengthDirection.SelectedIndex);
            WidthDirection = GetDirection(CmbWidthDirection.SelectedIndex);
            DialogResult = true;
        }

        private static string FormatDimension(double valueMm)
        {
            return valueMm > 0
                ? Math.Round(valueMm, 0).ToString(CultureInfo.InvariantCulture)
                : "";
        }

        private static bool TryParsePositiveDouble(string text, out double value)
        {
            string normalized = (text ?? "").Trim().Replace(',', '.');
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value) && value > 0.0;
        }

        private static int GetDirection(int selectedIndex)
        {
            if (selectedIndex == 1)
                return -1;
            if (selectedIndex == 2)
                return 1;
            return 0;
        }
    }
}