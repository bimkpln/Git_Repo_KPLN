using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_Ribbon.Common.Lists;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class FormListRenumber : Window
    {
        private readonly List<Parameter> _tBParams;
        public bool IsRun = false;

        public FormListRenumber(ParameterSet TBS)
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
            ExpBlocker(ExpClearRenumb, false);

            //cmbUniCode.ItemsSource = UniCodesCollection.CorretcUniCodes;
        }

        private void UsePrefix(object sender, RoutedEventArgs e)
        {
            ExpBlocker(ExpUnicodes, false);
            ExpBlocker(ExpPrefixes, true);
            ExpBlocker(ExpParamRefresh, false);
            ExpBlocker(ExpClearRenumb, false);
        }

        private void RefreshParam(object sender, RoutedEventArgs e)
        {
            ExpBlocker(ExpUnicodes, false);
            ExpBlocker(ExpPrefixes, false);
            ExpBlocker(ExpParamRefresh, true);
            ExpBlocker(ExpClearRenumb, false);

            cmbParam2.ItemsSource = _tBParams;
        }

        private void UseClearRenumb(object sender, RoutedEventArgs e)
        {
            ExpBlocker(ExpUnicodes, false);
            ExpBlocker(ExpPrefixes, false);
            ExpBlocker(ExpParamRefresh, false);
            ExpBlocker(ExpClearRenumb, true);
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

        private void Prefixes_Renumb(object sender, EventArgs e)
        {
            if ((bool)Prefixes_IsRenumbering.IsChecked)
                Prefixes_StartNumberTBox.IsEnabled = true;
            else
                Prefixes_StartNumberTBox.IsEnabled = false;
        }

        private void Prefixes_ParamUbpdate(object sender, EventArgs e)
        {
            if ((bool)Prefixes_IsRenumbering.IsChecked)
                Prefixes_StartNumberTBox.IsEnabled = true;
            else
                Prefixes_StartNumberTBox.IsEnabled = false;
        }

        private void UseUNICode_UniCouneter(object sender, EventArgs e)
        {
            if ((bool)UseUNICode_IsStrongUniCounter.IsChecked)
                UseUNICode_NumberOfUniCode.IsEnabled = true;
            else
                UseUNICode_NumberOfUniCode.IsEnabled = false;
        }

        private void ExpBlocker(Expander exp, bool rule)
        {
            exp.IsExpanded = rule;
            exp.IsEnabled = rule;
        }
    }
}
