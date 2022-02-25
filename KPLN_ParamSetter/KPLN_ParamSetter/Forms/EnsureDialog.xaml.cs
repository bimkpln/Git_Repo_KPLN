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

namespace KPLN_ParamManager.Forms
{
    public partial class EnsureDialog : Window
    {
        public string SIcon { get; set; }
        public string Header { get; set; }
        public string MainContent { get; set; }
        public static bool Commited = false;
        public EnsureDialog(Window parent, string sIcon, string header, string mainContent, bool cancelIsEnabled = true)
        {
            if(!cancelIsEnabled)
            { btnCancel.Visibility = Visibility.Collapsed; }
            SIcon = sIcon;
            Header = header;
            MainContent = mainContent;
            Commited = false;
            Owner = parent;
            InitializeComponent();
            DataContext = this;
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            Commited = true;
            Close();
        }
        private void OnDeny(object sender, RoutedEventArgs e)
        {
            Commited = false;
            Close();
        }
    }
}
