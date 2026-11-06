using KPLN_OpeningHoleManager.Core;
using System.Windows;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class OpeningHistoryFilterWindow : Window
    {
        public OpeningHistoryFilterSettings Settings { get; }

        public OpeningHistoryFilterWindow(OpeningHistoryFilterSettings settings)
        {
            Settings = settings.Copy();
            InitializeComponent();
            DataContext = Settings;
        }

        private void Reset_Click(object sender, RoutedEventArgs e) => Settings.Reset();
        private void Apply_Click(object sender, RoutedEventArgs e) => DialogResult = true;
    }
}
