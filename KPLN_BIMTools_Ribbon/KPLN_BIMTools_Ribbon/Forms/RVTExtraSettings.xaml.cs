using Autodesk.Revit.DB;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class RVTExtraSettings : UserControl
    {
        public DBRVTConfigData CurrentDBRSConfigData { get; set; }

        public RVTExtraSettings(DBRVTConfigData сurrentDBRSConfigData)
        {
            CurrentDBRSConfigData = сurrentDBRSConfigData;
            InitializeComponent();

            DataContext = CurrentDBRSConfigData;

            if (CurrentDBRSConfigData.MaxBackup != -1)
                MaxBackUpTBox.IsEnabled = true;

            if (CurrentDBRSConfigData.NameChangeFind != "🔐")
                NameChangeFindTBox.IsEnabled = true;

            if (CurrentDBRSConfigData.NameChangeSet != "🔐")
                NameChangeSetTBox.IsEnabled = true;
        }

        private void MaxBackupTBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string input = (sender as TextBox).Text;
            if (!double.TryParse(input, out double _))
            {
                UserDialog userDialog = new UserDialog("Ошибка", "Для количества резервных копий можно вводить только числа! Если не исправишь - будет значение по умолчанию = 10");
                userDialog.ShowDialog();
                CurrentDBRSConfigData.MaxBackup = 10;
            }
        }

        private void NameChangeFindTBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string nameChangeFind = (sender as TextBox).Text;
            if (!string.IsNullOrEmpty(nameChangeFind))
            {
                NameChangeSetTBox.IsEnabled = true;
                if(CurrentDBRSConfigData.NameChangeSet == "🔐")
                    CurrentDBRSConfigData.NameChangeSet = string.Empty;
            }
            else
            {
                NameChangeSetTBox.IsEnabled = false;
                CurrentDBRSConfigData.NameChangeSet = "🔐";
            }
        }
    }
}
