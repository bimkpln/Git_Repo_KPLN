using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

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

            DepartmentComboBox.Items.Clear();
            foreach (var dept in _departments.Where(d => d != departmentName))
            {
                DepartmentComboBox.Items.Add(new ComboBoxItem { Content = dept });
            }
            DepartmentComboBox.SelectedIndex = 0;
        }

        // Метод загрузки настроек
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
                MessageBox.Show($"Ошибка при чтении файла настроек:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return null; 
            }
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
                    writer.WriteLine((DepartmentComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано");
                    writer.WriteLine((HoleTypeComboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "Не выбрано");
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
                DepartmentComboBox.SelectedIndex = 0;
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