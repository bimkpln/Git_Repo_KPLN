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
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KPLN_NavisWorksReports.Forms
{
    /// <summary>
    /// Логика взаимодействия для TextInput.xaml
    /// </summary>
    public partial class TextInput : Window
    {
        public TextInput(Window parent, string header)
        {
            Owner = parent;
            TextInputDialog.Value = null;
            InitializeComponent();
            tbHeader.Text = header;
        }

        private void OnBtnApply(object sender, RoutedEventArgs e)
        {
            if ((tbox).Text.Length != 0)
            {
                TextInputDialog.Value = (tbox).Text;
                Close();
            }
            else
            {
                tbHeader.Foreground = Brushes.Red;
            }
        }
    }
}
