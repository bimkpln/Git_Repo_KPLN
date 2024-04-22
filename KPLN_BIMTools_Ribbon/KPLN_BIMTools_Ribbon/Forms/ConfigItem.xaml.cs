using Autodesk.Revit.DB;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Microsoft.Win32;
using NLog;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class ConfigItem : Window
    {
        private readonly Logger _logger;
        private readonly RevitDocExchangestDbService _revitDocExchangestDbService;
        private readonly SQLiteService _sqliteService;
        private readonly DBProject _project;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;
        private readonly List<CheckBox> _checkBoxList = new List<CheckBox>();
        private readonly List<FileEntity> _fileEntityList = new List<FileEntity>();
        private bool _canRunByName;
        private bool _canRunByPathTo;

        public ConfigItem(Logger logger, RevitDocExchangestDbService revitDocExchangestDbService, SQLiteService sqliteService, DBProject project, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            _logger = logger;
            _revitDocExchangestDbService = revitDocExchangestDbService;
            _sqliteService = sqliteService;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;

            InitializeComponent();
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            
            SetExtraSettings();

            DataContext = this;
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        /// <summary>
        /// Имя текущего конфига
        /// </summary>
        internal string SettingName { get; private set; }

        /// <summary>
        /// Общий путь для сохранения
        /// </summary>
        internal string SharedPathTo { get; private set; }

        /// <summary>
        /// Ссылка на текущий конфиг
        /// </summary>
        internal DBRevitDocExchanges CurrentDBRevitDocExchanges { get; private set; }

        /// <summary>
        /// Ссылка коллекцию настроек в конфиге
        /// </summary>
        internal List<DBConfigEntity> DBConfigEntityColl { get; private set; }

        /// <summary>
        /// Доп. настройки - UserControl
        /// </summary>
        public UserControl SelectedConfig { get; private set; }

        /// <summary>
        /// Динамическое создание блока доп. параметров (добавление в окно wpf), в зависимости от типа экспорта
        /// </summary>
        private void SetExtraSettings()
        {
            switch (_revitDocExchangeEnum)
            {
                case (RevitDocExchangeEnum.Navisworks):
                    SelectedConfig = new NWExtraSettings(new DBNWConfigData()
                    {
                        FacetingFactor = 1,
                        ConvertElementProperties = false,
                        ExportLinks = false,
                        FindMissingMaterials = false,
                        ExportScope = NavisworksExportScope.View,
                        DivideFileIntoLevels = true,
                        ExportRoomGeometry = true,
                        ViewName = "Navisworks",
                        WorksetToCloseNamesStartWith = "00_",
                        NavisDocPostfix = string.Empty,
                    });
                    break;
                case (RevitDocExchangeEnum.RevitServer):
                    break;
            }
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        #region Общие параметры
        private void ItemNameChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            SettingName = textBox.Text;

            if (!string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 5)
                _canRunByName = true;
            else 
                _canRunByName = false;

            BtnEnableSwitch();
        }

        private void PathToChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            SharedPathTo = textBox.Text;

            if (!string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 10)
                _canRunByPathTo = true;
            else 
                _canRunByPathTo = false;

            BtnEnableSwitch();
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            if (_canRunByName && _canRunByPathTo)
                btnOk.IsEnabled = true;
            else
                btnOk.IsEnabled = false;
        }
        #endregion

        #region Добавление файлов
        private void OnWindowsAddFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Revit Files (*.rvt)|*.rvt",
                InitialDirectory = @"Y:\",
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    AddCheckBox(new FileEntity(fileName, filePath));
                }
            }
        }

        private void OnBtnHandleAddFile(object sender, RoutedEventArgs e)
        {
            UserStringInput userStringInput = new UserStringInput();
            userStringInput.ShowDialog();
            if (userStringInput.IsRun)
            {
                AddCheckBox(new FileEntity(userStringInput.UserInputName, userStringInput.UserInputPath));
            }
        }

        private void AddCheckBox(FileEntity fileEntity)
        {
            btnRemoveFile.IsEnabled = true;

            CheckBox checkBox = new CheckBox
            {
                IsChecked = true,
                Content = new TextBlock
                {
                    Text = fileEntity.Name,
                    TextWrapping = TextWrapping.NoWrap,
                    ToolTip = fileEntity.Path,
                    Margin = new Thickness(5, 0, 0, 5),
                },
            };
            checkBox.Checked += CheckBox_Checked;
            checkBox.Unchecked += CheckBox_Unchecked;

            _checkBoxList.Add(checkBox);
            fileWrapPanel.Children.Add(checkBox);

            _fileEntityList.Add(fileEntity);
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            btnRemoveFile.IsEnabled = true;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_checkBoxList.Count(checkBox => checkBox.IsChecked == true) > 0)
                btnRemoveFile.IsEnabled = true;
            else
                btnRemoveFile.IsEnabled = false;
        }

        private void OnBtnRemoveFile(object sender, RoutedEventArgs e)
        {
            List<CheckBox> checkedCheckBoxes = new List<CheckBox>();

            foreach (CheckBox checkBox in _checkBoxList)
            {
                if (checkBox.IsChecked == true)
                {
                    checkedCheckBoxes.Add(checkBox);
                }
            }

            foreach (CheckBox checkBox in checkedCheckBoxes)
            {
                int index = _checkBoxList.IndexOf(checkBox);
                
                _checkBoxList.Remove(checkBox);
                fileWrapPanel.Children.Remove(checkBox);

                _fileEntityList.RemoveAt(index);
            }
        }
        #endregion

        #region Управление сохранением конфига
        private void OnBtnCancelClick(object sender, RoutedEventArgs e)
        {
            IsRun = false;
            this.Close();
        }

        private void OnBtnOkClick(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            
            // Создаю БД для записи
            Task createDB = Task.Run(() =>
            {
                _sqliteService.CreateDbFile();
            });

            CurrentDBRevitDocExchanges = new DBRevitDocExchanges
            {
                ProjectId = _project.Id,
                RevitDocExchangeType = _revitDocExchangeEnum.ToString(),
                SettingName = this.SettingName,
                SettingDBFilePath = _sqliteService.CurrentDBFullPath,
            };
            Task createDocExchanges = Task.Run(() =>
            {
                _revitDocExchangestDbService.CreateDBRevitDocExchanges(CurrentDBRevitDocExchanges);
            });

            //Создание экземпляра класса на основе введенных данных
            createDB.Wait();
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    NWExtraSettings extraSettings = (NWExtraSettings)SelectedConfig;
                    DBNWConfigData sharedDBNWConfigData = extraSettings.CurrentDBNWConfigData;

                    IEnumerable<DBNWConfigData> dBNWConfigDatas = _fileEntityList
                        .Select(fe => new DBNWConfigData(fe.Name, fe.Path, SharedPathTo).MergeWithDBConfigEntity(sharedDBNWConfigData));
                    _sqliteService.PostConfigItems_ByNWConfigs(dBNWConfigDatas);
                    break;
                case RevitDocExchangeEnum.RevitServer:
                    IEnumerable<DBRSConfigData> dBRSConfigDatas = _fileEntityList.Select(fe => new DBRSConfigData(fe.Name, fe.Path, SharedPathTo));
                    _sqliteService.PostConfigItems_ByRSConfigs(dBRSConfigDatas);
                    break;
            }

            createDocExchanges.Wait();
            this.Close();
        }
        #endregion
    }
}
