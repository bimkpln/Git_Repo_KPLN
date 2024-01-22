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

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class Example : Window
    {
        public Example()
        {
            InitializeComponent();
        }
        private void OnOk(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Renumb(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ListRenumb(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
