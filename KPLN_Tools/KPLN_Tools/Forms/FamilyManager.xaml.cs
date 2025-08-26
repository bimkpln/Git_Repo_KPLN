using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class FamilyManager : UserControl
    {
        string _currentSubDep;

        public FamilyManager(string currentStr)
        {
            InitializeComponent();

            _currentSubDep = currentStr;
            if (RunCurrentValue != null) RunCurrentValue.Text = string.IsNullOrWhiteSpace(_currentSubDep) ? "Не определён" : _currentSubDep;
            if (_currentSubDep != "BIM") BtnSettings.IsEnabled = false;


        }

        private void BtnSettings_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            TaskDialog.Show($"Информация ({_currentSubDep})", "Пока тут ничего нет :)");
        }
    }
}
