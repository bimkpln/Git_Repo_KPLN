using KPLN_Classificator.Forms.ViewModels;
using KPLN_Classificator.Utils;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_Classificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для ParameterSelectorForm.xaml
    /// </summary>
    public partial class ParameterSelectorForm : Window
    {
        public List<MyParameter> mparams;
        public BaseViewModel someModel;

        public ParameterSelectorForm(List<MyParameter> mparams, BaseViewModel someModel)
        {
            InitializeComponent();
            this.someModel = someModel;
            this.mparams = mparams;
            this.Collection.ItemsSource = this.mparams;
        }

        private void Accept_ParamName_Click(object sender, RoutedEventArgs e)
        {
            MyParameter parameter = (MyParameter)Collection.SelectedItem;
            if (someModel is ParamNameItem)
            {
                (someModel as ParamNameItem).paramName = parameter.Name;
            }
            else if (someModel is ParamValueItem)
            {
                ParamValueItem valueItem = someModel as ParamValueItem;
                var regex = new Regex(Regex.Escape("[]"));
                valueItem.paramValue = regex.Replace(valueItem.paramValue, string.Format("[{0}]", parameter.Name), 1);
            }
            this.Close();
        }
    }
}
