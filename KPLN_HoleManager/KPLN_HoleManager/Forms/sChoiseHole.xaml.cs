using System.Windows;
using Autodesk.Revit.DB;

namespace KPLN_HoleManager.Forms
{
    public partial class sChoiseHole : Window
    {
        private readonly Element _selectedElement;

        public sChoiseHole(string userFullName, string departmentName, Element selectedElement)
        {
            InitializeComponent();
            _selectedElement = selectedElement;

            // Устанавливаем DataContext для привязки данных в XAML
            DataContext = new HoleSelectionViewModel(userFullName, departmentName, selectedElement);
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}