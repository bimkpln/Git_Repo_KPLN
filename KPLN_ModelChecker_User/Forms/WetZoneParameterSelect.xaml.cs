using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class WetZoneParameterWindow : Window
    {
        public string SelectedParameter { get; private set; }

        public WetZoneParameterWindow(IEnumerable<Element> rooms)
        {
            InitializeComponent();

            RunRoomCount.Text = rooms.Count().ToString();
            List<string> fixedParameterNames = new List<string>
            {
                "Имя",
                "Назначение",
                "Комментарии"
            };

            ComboBoxParameters.ItemsSource = fixedParameterNames;
            ComboBoxParameters.SelectedItem = "Имя";
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxParameters.SelectedItem != null)
            {
                SelectedParameter = ComboBoxParameters.SelectedItem.ToString();
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите параметр.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
