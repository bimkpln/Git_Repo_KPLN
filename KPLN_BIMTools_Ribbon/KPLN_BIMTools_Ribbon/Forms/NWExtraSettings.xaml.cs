using Autodesk.Revit.DB;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class NWExtraSettings : UserControl
    {
        public DBNWConfigData CurrentDBNWConfigData { get; set; }

        public NWExtraSettings(DBNWConfigData currentDBNWConfigData)
        {
            CurrentDBNWConfigData = currentDBNWConfigData;
            InitializeComponent();

            ExportScopeComboBox.ItemsSource = new ObservableCollection<string>
            {
                NavisworksExportScope.Model.ToString(),
                NavisworksExportScope.View.ToString(),
                NavisworksExportScope.SelectedElements.ToString(),
            };
            ExportScopeComboBox.SelectedIndex = (int)CurrentDBNWConfigData.ExportScope;

            DataContext = CurrentDBNWConfigData;
        }

        private void ExportScope_Selected(object sender, System.Windows.RoutedEventArgs e) =>
            CurrentDBNWConfigData.ExportScope = (NavisworksExportScope)(sender as ComboBox).SelectedIndex;

        private void FacTBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string input = (sender as TextBox).Text;
            if (!double.TryParse(input, out double _))
            {
                UserDialog userDialog = new UserDialog("Ошибка", "Для коэффициента фасетизации можно вводить только числа! Если не исправишь - будет значение по умолчанию = 1");
                userDialog.ShowDialog();
                CurrentDBNWConfigData.FacetingFactor = 1.0;
            }
        }
    }
}
