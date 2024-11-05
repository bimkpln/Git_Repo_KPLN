using Autodesk.Revit.UI;
using System.Windows;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using System.Collections.Generic;



namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowMultipleLoadParameters : Window
    {
        UIApplication uiapp;
        public string activeFamilyName;

        public batchAddingParametersWindowMultipleLoadParameters(UIApplication uiapp)
        {        
            InitializeComponent();
            this.uiapp = uiapp;
           
        }

        // Разблокирование элементов интерфейса
        public void OpenInterfaceElements(string jsonFileSettingPath)
        {           
            TB_pathJson.Text = jsonFileSettingPath;
            SP_pathJsonFields.IsEnabled = false;
            SP_allPanelFamilyFields.IsEnabled = true;
            B_addParamsInFamily.IsEnabled = true;
            B_addFamilyInInterface.IsEnabled = true;
        }

        //// XAML. Открытие JSON-файла преднастроек
        private void OpenJSON_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Преднастройка (*.json)|*.json";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                string jsonFileSettingPath = openFileDialog.FileName;

                string jsonContent = System.IO.File.ReadAllText(jsonFileSettingPath);
                dynamic jsonFile = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);

                if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["pathFile"] != null && item["groupParameter"] != null && item["nameParameter"] != null && item["instance"] != null
                        && item["grouping"] != null && item["parameterValue"] != null && item["parameterValueDataType"] != null))
                {
                    paramTypeStatus.Content = "Преднастройка параметров с ФОП";
                    OpenInterfaceElements(jsonFileSettingPath);

                }
                else if (jsonFile is JArray && ((JArray)jsonFile).All(item =>
                        item["NE"] != null && item["quantity"] != null && item["parameterName"] != null && item["instance"] != null && item["categoryType"] != null
                        && item["dataType"] != null && item["grouping"] != null && item["parameterValue"] != null && item["comment"] != null))
                {
                    paramTypeStatus.Content = "Преднастройка с параметрами семейства";
                    OpenInterfaceElements(jsonFileSettingPath);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Ваш JSON-файл не является файлом преднастроек или повреждён. " +
                        "Пожалуйста, выберите другой файл.", "Ошибка чтения JSON-файла.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        //// XAML. Открытие файла семейства
        private void OpenFamily_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Семейство (*.rfa)|*.rfa";
            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                TB_familyPath.Text = openFileDialog.FileName;
            } 
        }

        //// XAML. Удалить поле с ссылкой на семейство
        private void DeleteFamilyField_Click(object sender, RoutedEventArgs e)
        {
            Button buttonDel = sender as Button;
            StackPanel panel = buttonDel.Parent as StackPanel;

            if (panel != null)
            {
                System.Windows.Controls.Panel parent = panel.Parent as System.Windows.Controls.Panel;
                if (parent != null)
                {
                    parent.Children.Remove(panel);
                }
            }
        }

        //// XAML. Добавление ссылки на семейство в интерфейс
        private void B_addFamilyInInterface_Click(object sender, RoutedEventArgs e)
        {
            StackPanel newStackPanel = new StackPanel
            {
                Tag = "uniqueFamilyField",
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Top,
                Height = 24,
                Margin = new Thickness(10, 10, 0, 0)
            };

            System.Windows.Controls.TextBox tbFamilyPath = new System.Windows.Controls.TextBox
            {
                IsReadOnly = true,
                Text = "Выберите семейство",
                Width = 815,
                Padding = new Thickness(4, 3, 0, 0),
                Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFF4FBB5"))
            };

            Button openButton = new Button
            {
                Content = "Открыть",
                Width = 75,
                Height = 24
            };

            openButton.Click += (s, ev) =>
            {
                Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
                openFileDialog.Filter = "Семейство (*.rfa)|*.rfa";
                bool? result = openFileDialog.ShowDialog();

                if (result == true)
                {
                    tbFamilyPath.Text = openFileDialog.FileName;
                }
            };

            Button deleteButton = new Button
            {
                Content = "X",
                Width = 25,
                Height = 24,
                Background = new SolidColorBrush(Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            deleteButton.Click += DeleteFamilyField_Click;

            newStackPanel.Children.Add(tbFamilyPath);
            newStackPanel.Children.Add(openButton);
            newStackPanel.Children.Add(deleteButton);

            SP_allPanelFamilyFields.Children.Add(newStackPanel);
        }

        //// XAML. Добавление параметров в семейство
        private void B_addParamsInFamily_Click(object sender, RoutedEventArgs e)
        {
            // Проверка на дубликаты ссылок семейств 

            // Проверка на пустые ссылки семейства 

            // Сам процесс

        }
    }
}
