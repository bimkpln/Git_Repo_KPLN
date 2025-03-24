using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace KPLN_HoleManager.Forms
{
    public partial class DockableManagerFormSettings : Window
    {
        // Переданные данные
        private string _userFullName;
        private string _departmentName;

        private readonly string[] _departments = { "Не выбрано", "АР", "КР", "ОВиК", "ВК", "ЭОМ", "СС" };
        private readonly static string _settingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Autodesk", "Revit", "holeManagerSetting.ini");

        public DockableManagerFormSettings(string userFullName, string departmentName)
        {
            InitializeComponent();

            _userFullName = userFullName;
            _departmentName = departmentName;

            // Задание от отдела
            if (_departmentName == "BIM")
            {
                DepartmentFromComboBox.IsEnabled = true;
                DepartmentFromComboBox.Foreground = new SolidColorBrush(Colors.Black);

                foreach (ComboBoxItem item in DepartmentFromComboBox.Items)
                {
                    if (item.Content.ToString() == "Не выбрано")
                    {
                        DepartmentFromComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
            else
            {
                foreach (ComboBoxItem item in DepartmentFromComboBox.Items)
                {
                    if (item.Content.ToString() == _departmentName)
                    {
                        DepartmentFromComboBox.SelectedItem = item;
                        break;
                    }
                }         
            }

            // Загрузка преднастроек при наличии
            List<string> setting = LoadSettings();
            if (setting != null)
            {
                foreach (ComboBoxItem item in DepartmentFromComboBox.Items)
                {
                    if (item.Content.ToString() == setting[2])
                    {
                        DepartmentFromComboBox.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in DepartmentInComboBox.Items)
                {
                    if (item.Content.ToString() == setting[3])
                    {
                        DepartmentInComboBox.SelectedItem = item;
                        break;
                    }
                }

                if (setting[4] == "SquareHole") setting[4] = "Прямоугольное";
                else if (setting[4] == "RoundHole") setting[4] = "Круглое";

                foreach (ComboBoxItem item in HoleTypeComboBox.Items)
                {
                    if (item.Content.ToString() == setting[4])
                    {
                        HoleTypeComboBox.SelectedItem = item;
                        break;
                    }
                }

                foreach (ComboBoxItem item in HoleIndentComboBox.Items)
                {
                    if (item.Content.ToString() == setting[5])
                    {
                        HoleIndentComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        // Метод получения преднастроек
        public static List<string> LoadSettings()
        {
            if (!File.Exists(_settingsPath))
            {
                return null; 
            }

            try
            {
                List<string> settings = File.ReadAllLines(_settingsPath).ToList();

                if (settings.Count < 5)
                {
                    return null; 
                }

                return settings;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении файла настроек:\n{ex.Message}", 
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return null; 
            }
        }

        // XAML. Поведение ComboBox DepartmentFromComboBox
        private void DepartmentFromComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DepartmentInComboBox == null) return;

            // Очищаем текущие элементы
            DepartmentInComboBox.Items.Clear();

            ComboBoxItem selectedItem = DepartmentFromComboBox.SelectedItem as ComboBoxItem;
            if (selectedItem == null) return;

            string selectedText = selectedItem.Content.ToString();

            // Добавляем "Не выбрано" в любом случае
            DepartmentInComboBox.Items.Add(new ComboBoxItem { Content = "Не выбрано" });

            // Если выбран один из "ОВиК", "ВК", "ЭОМ" или "СС"
            if (selectedText == "ОВиК" || selectedText == "ВК" || selectedText == "ЭОМ" || selectedText == "СС")
            {
                // Добавляем только "АР" и "КР"
                DepartmentInComboBox.Items.Add(new ComboBoxItem { Content = "АР" });
                DepartmentInComboBox.Items.Add(new ComboBoxItem { Content = "КР" });
            }
            else
            {
                // Добавляем все возможные элементы, кроме текущего выбранного
                string[] allDepartments = { "АР", "КР", "ОВиК", "ВК", "ЭОМ", "СС" };

                foreach (var department in allDepartments)
                {
                    if (department != selectedText)
                    {
                        DepartmentInComboBox.Items.Add(new ComboBoxItem { Content = department });
                    }
                }
            }

            DepartmentInComboBox.SelectedIndex = 0;
        }


        // XAML. Сохранение файла настроек
        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(_settingsPath, false))
                {
                    writer.WriteLine(_userFullName);
                    writer.WriteLine(_departmentName);
                    writer.WriteLine((DepartmentFromComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано");
                    writer.WriteLine((DepartmentInComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано");
                    string holeType = (HoleTypeComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано";

                    if (holeType == "Прямоугольное") holeType = "SquareHole";
                    else if (holeType == "Круглое") holeType = "RoundHole";
                    writer.WriteLine(holeType);

                    writer.WriteLine((HoleIndentComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано");

                    MessageBox.Show("Настройки сохранены.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении настроек:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // XAML. Удаление файла настроек
        private void CleanSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    File.Delete(_settingsPath);
                }

                // Сбрасываем все ComboBox к начальному значению
                if (_departmentName == "BIM")
                {
                    DepartmentFromComboBox.SelectedIndex = 0;
                }

                DepartmentInComboBox.SelectedIndex = 0;
                HoleTypeComboBox.SelectedIndex = 0;
                HoleIndentComboBox.SelectedIndex = 0;

                MessageBox.Show("Настройки сброшены.", "Очистка", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сбросе настроек:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // XAML. Закрытие окна
        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}