using Autodesk.Revit.DB;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.ObjectModel;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class IFCExtraSettings : UserControl
    {
        public DBIFCConfigData CurrentDBIFCConfigData { get; set; }

        public IFCExtraSettings(DBIFCConfigData currentDBIFCConfigData)
        {
            CurrentDBIFCConfigData = currentDBIFCConfigData;
            InitializeComponent();

            FileVersionComboBox.ItemsSource = new ObservableCollection<string>(Enum.GetNames(typeof(IFCVersion)));
            FileVersionComboBox.SelectedItem = CurrentDBIFCConfigData.FileVersion.ToString();

            DataContext = CurrentDBIFCConfigData;
        }

        private void FileVersion_Selected(object sender, System.Windows.RoutedEventArgs e)
        {
            string selectedValue = (sender as ComboBox)?.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(selectedValue))
                return;

            CurrentDBIFCConfigData.FileVersion = (IFCVersion)Enum.Parse(typeof(IFCVersion), selectedValue);
        }

        private void SpaceBoundaryLevel_TextChanged(object sender, TextChangedEventArgs e)
        {
            string input = (sender as TextBox)?.Text;
            if (!int.TryParse(input, out int value) || value < 0 || value > 2)
            {
                UserDialog userDialog = new UserDialog(
                    "Ошибка",
                    "Для уровня границ пространств можно вводить только 0, 1 или 2. Если не исправишь - будет значение по умолчанию = 0");

                userDialog.ShowDialog();
                CurrentDBIFCConfigData.SpaceBoundaryLevel = 0;
            }
        }
    }
}