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

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для KR_WSofLinks.xaml
    /// </summary>
    public partial class KR_WSofLinks : Window
    {

        public string WorksetName { get; private set; }
        public bool WorksetOpenClose { get; private set; }

        public KR_WSofLinks()
        {
            InitializeComponent();
            tbWorksetName.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            //передаем имя введенного РН
            WorksetName = tbWorksetName.Text.Trim();            
            //передаем режим работы (Открыть/Закрыть)
            WorksetOpenClose = chkOpenClose.IsChecked == true;      
            DialogResult = !string.IsNullOrEmpty(WorksetName);
        }

    }
}
