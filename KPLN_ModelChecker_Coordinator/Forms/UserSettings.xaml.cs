using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_ModelChecker_Coordinator.Forms
{
    /// <summary>
    /// Логика взаимодействия для UserSettings.xaml
    /// </summary>
    public partial class UserSettings : Window
    {
        public UserSettings()
        {
            InitializeComponent();
#if Revit2020
            Owner = ModuleData.RevitWindow;
#endif
#if Revit2018
            WindowInteropHelper helper = new WindowInteropHelper(this);
            helper.Owner = ModuleData.MainWindowHandle;
#endif
            this.chbx_dialogs.IsChecked = ModuleData.up_close_dialogs;
            this.chbx_enter.IsChecked = ModuleData.up_send_enter;
            this.chbx_telegram.IsChecked = ModuleData.up_notify_in_tg;
        }

        private void OnApply(object sender, RoutedEventArgs e)
        {
            ModuleData.up_close_dialogs = (bool)this.chbx_dialogs.IsChecked;
            ModuleData.up_send_enter = (bool)this.chbx_enter.IsChecked;
            ModuleData.up_notify_in_tg = (bool)this.chbx_telegram.IsChecked;
            Close();
        }
    }
}
