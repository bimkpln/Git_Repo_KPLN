﻿using System.Collections.Generic;
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
        private string _sendingDepartmentHoleName;
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

        // Настраиваем ComboBox
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

        // XAML. Выпадающий список с отделамми
        private void DepartmentComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                string selectedDepartment = selectedItem.Content.ToString();
                DepartmentComboBox.Foreground = Brushes.Black;

                // Определяем доступные отделы без выбранного
                List<string> availableDepartments = new List<string> { "АР", "КР", "ИОС" };
                availableDepartments.Remove(selectedDepartment);

                // Очищаем SendingDepartmentComboBox и заполняем его новыми значениями
                SendingDepartmentComboBox.Items.Clear();
                foreach (string department in availableDepartments)
                {
                    SendingDepartmentComboBox.Items.Add(new ComboBoxItem { Content = department });
                }

                // Определяем, какой элемент должен быть выбран по умолчанию
                string defaultSendingDepartment = selectedDepartment == "ИОС" ? "АР" : "ИОС";

                // Устанавливаем нужный элемент в SendingDepartmentComboBox
                foreach (ComboBoxItem item in SendingDepartmentComboBox.Items)
                {
                    if (item.Content.ToString() == defaultSendingDepartment)
                    {
                        SendingDepartmentComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
        }

        // XAML. Кнопка с прямоугольным отверстием
        private void SquareHoleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                _departmentHoleName = selectedItem.Content.ToString();
                _sendingDepartmentHoleName = (SendingDepartmentComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                _holeTypeName = "SquareHole";

                RunPlaceHoleCommand();
            }
        }

        // XAML. Кнопка с круглым отверстием
        private void RoundHoleButton_Click(object sender, RoutedEventArgs e)
        {
            if (DepartmentComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                _departmentHoleName = selectedItem.Content.ToString();
                _sendingDepartmentHoleName = (SendingDepartmentComboBox.SelectedItem as ComboBoxItem)?.Content.ToString();
                _holeTypeName = "RoundHole";

                RunPlaceHoleCommand();
            }
        }

        // XAML. Закрытие окна
        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // Вызов команды _ExternalEventHandler для срабатывание Execute
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
                PlaceHoleOnWallCommand.Execute(app, _userFullName, _departmentName, _selectedElement, _departmentHoleName, _sendingDepartmentHoleName, _holeTypeName);
            });
        }
    }
}