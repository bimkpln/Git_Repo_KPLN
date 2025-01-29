using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;

namespace KPLN_HoleManager.Forms
{
    public partial class sChoiseHole : Window
    {
        private readonly Element _selectedElement;

        private readonly string _userFullName;
        private readonly string _departmentName;

        private string _departmentHoleName;
        private string _holeTypeName;

        public sChoiseHole(string userFullName, string departmentName, Element selectedElement)
        {
            InitializeComponent();

            // Данные для дальнейшей передачи
            _selectedElement = selectedElement;            
            _userFullName = userFullName;
            _departmentName = departmentName;

            // Устанавливаем DataContext для привязки данных в XAML
            DataContext = new HoleSelectionViewModel(userFullName, departmentName, selectedElement);

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

                MessageBox.Show($"Выбранный элемент: {_selectedElement}\n" +
                    $"Пользователь: {_userFullName}\n" +
                    $"Отдел: {_departmentName}\n" +
                    $"Отдел дырки: {_departmentHoleName}\n" +
                    $"Тип дырки: {_holeTypeName}\n", 
                    "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Кнопка с круглым отверстием
        private void RoundHoleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                _departmentHoleName = selectedItem.Content.ToString();
                _holeTypeName = "RoundHole";

                MessageBox.Show($"Выбранный элемент: {_selectedElement}\n" +
                    $"Пользователь: {_userFullName}\n" +
                    $"Отдел: {_departmentName}\n" +
                    $"Отдел дырки: {_departmentHoleName}\n" +
                    $"Тип дырки: {_holeTypeName}\n",
                    "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        // Закрытие окна
        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}