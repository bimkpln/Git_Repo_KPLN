using System.Windows;
using System.Windows.Controls;

namespace KPLN_HoleManager.Forms
{
    public partial class sChoiseHoleIndentation : Window
    {
        public double SelectedOffset { get; private set; } = 0;

        public sChoiseHoleIndentation()
        {
            InitializeComponent();
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
        }

        private void OffsetButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && double.TryParse(button.Content.ToString().Replace(" мм", ""), out double offset))
            {
                SelectedOffset = offset;
                DialogResult = true;
                Close();
            }
        }
    }
}
