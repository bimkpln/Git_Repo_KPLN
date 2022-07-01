using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using KPLN_Views_Ribbon.Common.Lists;

namespace KPLN_Views_Ribbon.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class FormListRename : Window
    {
        public bool IsRun = false;
        private List<Parameter> _tBParams;

        public FormListRename(ParameterSet TBS)
        {
            List<Parameter> mixedParamList = new List<Parameter>();
            foreach (Parameter curParam in TBS)
            {
                mixedParamList.Add(curParam);
            }
            _tBParams = mixedParamList.OrderBy(p => p.Definition.Name).ToList();
            InitializeComponent();
        }
        
        private void UseUNICode(object sender, RoutedEventArgs e)
        {

            ExpBlocker(ExpUnicodes, true);
            ExpBlocker(ExpPrefixes, false);
            ExpBlocker(ExpParamRefresh, false);

            cmbUniCode.ItemsSource = new UniCodesCollection().UniCodes;
        }
        
        private void UsePrefix(object sender, RoutedEventArgs e)
        {
            ExpBlocker(ExpUnicodes, false);
            ExpBlocker(ExpPrefixes, true);
            ExpBlocker(ExpParamRefresh, false);
        }
        
        private void RefreshParam(object sender, RoutedEventArgs e)
        {
            ExpBlocker(ExpUnicodes, false);
            ExpBlocker(ExpPrefixes, false);
            ExpBlocker(ExpParamRefresh, true);
            
            cmbParam.ItemsSource = _tBParams;
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            Close();
        }
        
        private void OnCancel(object sender, RoutedEventArgs e)
        {
            Close();
        }
        
        private void OnInfo(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Renumb(object sender, EventArgs e)
        {
            if ((bool)isRenumbering.IsChecked)
            {
                strNumTextBox.IsEnabled = true;
            }
            else
            {
                strNumTextBox.IsEnabled = false;
            }
        }

        private void UniCouneter(object sender, EventArgs e)
        {
            if ((bool)isStrongUniCounter.IsChecked)
            {
                strNumUniCode.IsEnabled = true;
            }
            else
            {
                strNumUniCode.IsEnabled = false;
            }
        }

        private void ExpBlocker(Expander exp, bool rule)
        {
            exp.IsExpanded = rule;
            exp.IsEnabled = rule;
        }
    }
}
