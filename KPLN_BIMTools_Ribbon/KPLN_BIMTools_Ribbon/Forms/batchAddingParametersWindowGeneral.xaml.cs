using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;


namespace KPLN_BIMTools_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для batchAddingParameters.xaml
    /// </summary>
    public partial class batchAddingParametersWindowGeneral : Window
    {
        public UIApplication uiapp;
        public Autodesk.Revit.ApplicationServices.Application revitApp;

        public string paramAction;
        public string jsonFileSettingPath;

        public string openFileDialogFilePath;
        Dictionary<String, List<ExternalDefinition>> generalParametersFileDict = new Dictionary<String, List<ExternalDefinition>>();


        public batchAddingParametersWindowGeneral(UIApplication uiapp, string paramAction, string jsonFileSettingPath)
        {
            InitializeComponent();

            this.uiapp = uiapp;
            this.paramAction = paramAction;
            this.jsonFileSettingPath = jsonFileSettingPath;

            revitApp = uiapp.Application;

            if (!string.IsNullOrEmpty(jsonFileSettingPath))
            {
                openFileDialogFilePath = revitApp.SharedParametersFilename;
                TB_filePath.Text = openFileDialogFilePath;
                // Реализовать загрузку параметров из json-файла
            }
            else
            {
                openFileDialogFilePath = revitApp.SharedParametersFilename;
                TB_filePath.Text = openFileDialogFilePath;

                HandlerGeneralParametersFile(openFileDialogFilePath);
            }        
        }

        // Обработчик ФОПа c обработкой ComboBox "Группы"
        public void HandlerGeneralParametersFile(string filePath)
        {
            generalParametersFileDict.Clear();
            revitApp.SharedParametersFilename = openFileDialogFilePath;

            try
            {
                DefinitionFile defFile = revitApp.OpenSharedParameterFile();
                if (defFile == null)
                {
                    System.Windows.Forms.MessageBox.Show($"{openFileDialogFilePath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    return;
                }

                foreach (DefinitionGroup group in defFile.Groups)
                {
                    List<ExternalDefinition> parametersList = new List<ExternalDefinition>();

                    foreach (ExternalDefinition definition in group.Definitions)
                    {
                        parametersList.Add(definition);
                    }

                    generalParametersFileDict[group.Name] = parametersList;
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"{openFileDialogFilePath}\n" +
                        "Пожалуйста, выберете другой ФОП.", "Ошибка чтения ФОП.", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                
                openFileDialogFilePath = "";
                TB_filePath.Text = openFileDialogFilePath;
            }

            UpdateComboBoxField_Group();      
        }

        // Обновления ComboBox "Группы"
        private void UpdateComboBoxField_Group()
        {
            foreach (var key in generalParametersFileDict.Keys)
            {
                CB_paramsGroup.Items.Add(key);
            }
        }

        // Обновления ComboBox "Параметры"
        private void UpdateComboBoxField_Param(string selectGroup)
        {
            CB_paramsName.Items.Clear();

            if (generalParametersFileDict.ContainsKey(selectGroup))
            {
                foreach (var param in generalParametersFileDict[selectGroup])
                {
                    CB_paramsName.Items.Add(param.Name);
                }
            }
        }

        // XAML. Открытие ФОПа
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();

            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(revitApp.SharedParametersFilename);

            if (openFileDialog.ShowDialog() == true)
            {
                openFileDialogFilePath = openFileDialog.FileName;
                TB_filePath.Text = openFileDialogFilePath;

                HandlerGeneralParametersFile(openFileDialogFilePath);
            }
        }

        // XAML. Отслеживание ComboBox - Группы
        private void OnGroupSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CB_paramsGroup.SelectedItem != null)
            {
                string selectGroup = CB_paramsGroup.SelectedItem.ToString();
                UpdateComboBoxField_Param(selectGroup);
            }    
        }

        // XAML. Удалить SP_panelParamFields
        private void RemovePanel(object sender, RoutedEventArgs e)
        {
            Button button = sender as Button;
            StackPanel panel = button.Parent as StackPanel;

            if (panel != null)
            {
                System.Windows.Controls.Panel parent = panel.Parent as System.Windows.Controls.Panel;
                if (parent != null)
                {
                    parent.Children.Remove(panel);
                }
            }
        }

        // XAML. Добавить новый параметр в панель
        private void AddPanelParamFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            StackPanel newPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(20, 0, 20, 12)
            };

            System.Windows.Controls.ComboBox cbParamsName = new System.Windows.Controls.ComboBox
            {
                Width = 490,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
            };

            System.Windows.Controls.ComboBox cbParamsType = new System.Windows.Controls.ComboBox
            {
                Width = 105,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0),
                SelectedIndex = 0
            };

            cbParamsType.Items.Add(new ComboBoxItem { Content = "Тип" });
            cbParamsType.Items.Add(new ComboBoxItem { Content = "Экземпляр" });

            System.Windows.Controls.ComboBox cbParamsGroup = new System.Windows.Controls.ComboBox
            {
                Width = 270,
                Height = 25,
                Padding = new Thickness(8, 4, 0, 0)
            };

            foreach (var key in generalParametersFileDict.Keys)
            {               
                cbParamsGroup.Items.Add(key);
            }

            cbParamsGroup.SelectionChanged += (s, ev) =>
            {
                string selectedGroup = cbParamsGroup.SelectedItem.ToString();
                cbParamsName.Items.Clear();

                if (generalParametersFileDict.ContainsKey(selectedGroup))
                {
                    foreach (var param in generalParametersFileDict[selectedGroup])
                    {
                        cbParamsName.Items.Add(param.Name);
                    }
                }
            };

            Button removeButton = new Button
            {
                Width = 30,
                Height = 25,
                Content = "X",
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(158, 3, 3)),
                Foreground = new SolidColorBrush(Colors.White)
            };

            removeButton.Click += (s, ev) =>
            {
                SP_allPanelParamsFields.Children.Remove(newPanel);
            };

            newPanel.Children.Add(cbParamsName);
            newPanel.Children.Add(cbParamsType);
            newPanel.Children.Add(cbParamsGroup);
            newPanel.Children.Add(removeButton);

            SP_allPanelParamsFields.Children.Add(newPanel);
        }

    }
}
