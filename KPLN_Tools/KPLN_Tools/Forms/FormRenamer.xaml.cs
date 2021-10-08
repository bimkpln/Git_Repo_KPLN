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
using Autodesk.Revit.DB;


namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormRenamer.xaml
    /// </summary>
    public partial class FormRenamer : Window
    {
        public string prfTxt;
        public string strNumb;
        List<Parameter> TBParams;

        public FormRenamer(ParameterSet TBS)
        {
            List<Parameter> mixedParamList = new List<Parameter>();
            foreach(Parameter curParam in TBS)
            {
                mixedParamList.Add(curParam);
            }
            TBParams = mixedParamList.OrderBy(p => p.Definition.Name).ToList();
            InitializeComponent();
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            prfTxt = prfTextBox.Text;
            strNumb = strNumTextBox.Text;            
            this.Close();
        }


        private void OnCancel(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Renumb (object sender, EventArgs e)
        {
            if((bool)isRenumbering.IsChecked)
            {
                strNumTextBox.IsEnabled = true;
            }
            else
            {
                strNumTextBox.IsEnabled = false;

            }
        }

        private void ListRenumb(object sender, EventArgs e) 
        {
            if ((bool)isChangingList.IsChecked)
            {
                cmbParam.IsEnabled = true;
                cmbParam.ItemsSource = TBParams;
            }
            else
            {
                cmbParam.IsEnabled = false;
            }
        }

    }
}
