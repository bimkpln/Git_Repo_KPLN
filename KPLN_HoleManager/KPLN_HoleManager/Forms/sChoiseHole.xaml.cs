using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_HoleManager.ExternalCommand;

namespace KPLN_HoleManager.Forms
{
    public partial class sChoiseHole : Window
    {
        private readonly UIApplication _uiApp;
        private readonly Element _selectedElement;

        private readonly string _userFullName;
        private readonly string _departmentName;

        private string _departmentHoleName;
        private string _holeTypeName;

        public sChoiseHole(UIApplication uiApp, Element selectedElement, string userFullName, string departmentName)
        {
            InitializeComponent();

            // Данные для дальнейшей передачи
            _uiApp = uiApp;
            _selectedElement = selectedElement;            
            _userFullName = userFullName;
            _departmentName = departmentName;

            // Устанавливаем DataContext для привязки данных в XAML
            DataContext = new HoleSelectionViewModel(selectedElement, userFullName, departmentName);

            // Настраиваем ComboBox
            SetDepartmentComboBox();
        }

        private void SetDepartmentComboBox()
        {
            if (_departmentName == "BIM")
            {
                // Если BIM, разрешаем редактирование и устанавливаем "АР" по умолчанию
                DepartmentComboBox.IsEnabled = true;
                DepartmentComboBox.SelectedIndex = 0;
                DepartmentComboBox.Foreground = Brushes.Black;
            }
            else
            {
                // Если не BIM, блокируем редактирование, устанавливаем департамент и делаем текст серым
                DepartmentComboBox.IsEnabled = false;
                var departmentItem = DepartmentComboBox.Items
                    .Cast<ComboBoxItem>()
                    .FirstOrDefault(item => item.Content.ToString() == _departmentName);

                if (departmentItem != null)
                {
                    DepartmentComboBox.SelectedItem = departmentItem;
                }
                DepartmentComboBox.Foreground = Brushes.Gray;
            }
        }

        private void DepartmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Меняем цвет текста, если BIM редактирует поле
            if (DepartmentComboBox.IsEnabled)
            {
                DepartmentComboBox.Foreground = Brushes.Black;
            }
        }

        // Кнопка с прямоугольным отверстием
        private void SquareHoleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                _departmentHoleName = selectedItem.Content.ToString();
                _holeTypeName = "SquareHole";

                RunPlaceHoleCommand();
            }
        }

        // Кнопка с круглым отверстием
        private void RoundHoleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                _departmentHoleName = selectedItem.Content.ToString();
                _holeTypeName = "RoundHole";

                RunPlaceHoleCommand();
            }
        }

        // Вызов команды
        private void RunPlaceHoleCommand()
        {
            this.Close();

            if (_uiApp == null)
            {
                TaskDialog.Show("Ошибка", "Не удалось получить доступ к Revit.");
                return;
            }
          
            // Вызываем команду с параметрами
            _ExternalEventHandler.Instance.Raise((app) =>
            {
                PlaceHoleOnWallCommand.Execute(app, _selectedElement, _departmentHoleName, _holeTypeName);
            });
        }


        // Закрытие окна
        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}