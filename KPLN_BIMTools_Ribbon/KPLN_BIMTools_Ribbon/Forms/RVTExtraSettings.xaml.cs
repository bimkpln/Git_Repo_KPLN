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

            MaxBackUpTBox.Text = CurrentDBRSConfigData.MaxBackup == -1
                ? "🔐" : CurrentDBRSConfigData.MaxBackup.ToString(CultureInfo.CurrentCulture);

            if (CurrentDBRSConfigData.NameChangeFind != "🔐")
                NameChangeFindTBox.IsEnabled = true;

            if (CurrentDBRSConfigData.NameChangeSet != "🔐")
                NameChangeSetTBox.IsEnabled = true;
        }

        public bool HasValidMaxBackup => MaxBackUpTBox.Text.Trim() == "🔐"
            || (int.TryParse(MaxBackUpTBox.Text, out int count) && count > 0);

        private void MaxBackupTBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string input = ((TextBox)sender).Text.Trim();
            if (input == "🔐")
                CurrentDBRSConfigData.MaxBackup = -1;
            else if (int.TryParse(input, out int count) && count > 0)
                CurrentDBRSConfigData.MaxBackup = count;
        }

        private void MaxBackupTBox_GotKeyboardFocus(object sender, System.Windows.Input.KeyboardFocusChangedEventArgs e) =>
            ((TextBox)sender).SelectAll();

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
