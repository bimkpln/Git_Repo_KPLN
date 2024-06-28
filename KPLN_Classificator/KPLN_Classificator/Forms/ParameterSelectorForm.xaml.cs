using KPLN_Classificator.Utils;
using System.Collections.Generic;
using System.Windows;

namespace KPLN_Classificator.Forms
{
    /// <summary>
    /// Логика взаимодействия для ParameterSelectorForm.xaml
    /// </summary>
    public partial class ParameterSelectorForm : Window
    {
        public List<MyParameter> mparams;
        public MyParameter SelectedMyParameter;

        public ParameterSelectorForm(List<MyParameter> mparams)
        {
            InitializeComponent();
            this.mparams = mparams;
            this.Collection.ItemsSource = this.mparams;
        }

        private void Accept_ParamName_Click(object sender, RoutedEventArgs e)
        {
            SelectedMyParameter = (MyParameter)Collection.SelectedItem;
            this.Close();
        }
    }
}
