using KPLN_Loader.Common;
using System.Windows;
using System.Windows.Controls;
using System.Linq;

using static KPLN_Loader.Preferences;

namespace KPLN_Loader.Forms
{
    /// <summary>
    /// Логика взаимодействия для WPFLogIn.xaml
    /// </summary>
    public partial class WPFLogIn : Window
    {
        public WPFLogIn()
        {
            //Owner = Preferences.RevitWindow;
            InitializeComponent();
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (this.tbxName.Text != "" && this.tbxFamily.Text != "" && this.cbxDepartment.SelectedIndex != -1) { this.btnApply.IsEnabled = true; }
            else { this.btnApply.IsEnabled = false; }
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.tbxName.Text != "" && this.tbxFamily.Text != "" && this.cbxDepartment.SelectedIndex != -1) { this.btnApply.IsEnabled = true; }
            else { this.btnApply.IsEnabled = false; }
        }

        private void OnClickApply(object sender, RoutedEventArgs e)
        {
            User = new SQLUserInfo(System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last(),
                this.tbxName.Text,
                this.tbxSurname.Text,
                this.tbxFamily.Text,
                "");
            
            User.Department = this.cbxDepartment.SelectedItem as SQLDepartmentInfo;
            
            try
            {
                Tools_SQL.CreateUser(User.SystemName,
                        User.Name,
                        User.Family,
                        User.Surname,
                        User.Department.Id);
            }
            catch (System.Exception) { }
            
            this.Close();
        }

        private void OnClickClose(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
