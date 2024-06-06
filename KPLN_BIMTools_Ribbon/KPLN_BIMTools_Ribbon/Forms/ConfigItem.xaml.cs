using Autodesk.Revit.DB;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Microsoft.Win32;
using NLog;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class ConfigItem : Window, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Кэширование пути для выбора файлов
        /// </summary>
        private static string _initialDirectoryForOpenFileDialog = @"Y:\";

        private readonly Logger _logger;
        private readonly RevitDocExchangestDbService _revitDocExchangestDbService;
        private readonly SQLiteService _sqliteService;
        private readonly DBProject _project;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;

        private string _settingName;
        private string _sharedPathTo;
        private bool _canRunByName;
        private bool _canRunByPathTo;

        /// <summary>
        /// Конструктор основной единицы отчета
        /// </summary>
        /// <param name="logger">Логгер</param>
        /// <param name="revitDocExchangestDbService">Текущий сервис работы с БД по отчетам из диспетчера</param>
        /// <param name="sqliteService">Текущий сервис работы с БД по отчетам из текущего окна</param>
        /// <param name="project">Ссылка на проект</param>
        /// <param name="revitDocExchangeEnum">Тип обмена</param>
        /// <param name="currentDBRevitDocExchanges">Ссылка на существующий конфиг</param>
        public ConfigItem(
            Logger logger,
            RevitDocExchangestDbService revitDocExchangestDbService,
            SQLiteService sqliteService,
            DBProject project,
            RevitDocExchangeEnum revitDocExchangeEnum,
            DBRevitDocExchanges currentDBRevitDocExchanges = null)
        {
            _logger = logger;
            _revitDocExchangestDbService = revitDocExchangestDbService;
            _sqliteService = sqliteService;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;
            CurrentDBRevitDocExchanges = currentDBRevitDocExchanges;

            string mainProjectPath = _project.MainPath;
            if (!string.IsNullOrEmpty(mainProjectPath) && Directory.Exists(mainProjectPath))
                _initialDirectoryForOpenFileDialog = mainProjectPath;

            InitializeComponent();
            PreviewKeyDown += new KeyEventHandler(HandleEsc);

            if (CurrentDBRevitDocExchanges == null)
                SetExtraSettings();
            else
            {
                // Добавляю общее имя конфига
                CurrentDBRevitDocExchanges = _revitDocExchangestDbService.GetDBRevitDocExchanges_ById(CurrentDBRevitDocExchanges.Id);
                SettingName = CurrentDBRevitDocExchanges.SettingName;

                // Проверяю на триггер копирования - базы данных не будут совпадать
                SQLiteService tempSqliteService = null;
                if (_sqliteService.CurrentDBFullPath != CurrentDBRevitDocExchanges.SettingDBFilePath)
                    tempSqliteService = new SQLiteService(_logger, CurrentDBRevitDocExchanges.SettingDBFilePath, _revitDocExchangeEnum);
                else 
                    tempSqliteService = _sqliteService;

                // Добавляю общие настройки конфига (которые дублируются по каждтому DBConfigEntity)
                IEnumerable<DBConfigEntity> dBConfigEntities = tempSqliteService.GetConfigItems();
                SetExtraSettings_DBConfigEntity(dBConfigEntities.FirstOrDefault());
                SharedPathTo = dBConfigEntities.FirstOrDefault().PathTo;

                // Добавляю список файлов
                FileEntitiesList = new ObservableCollection<FileEntity>(dBConfigEntities.Select(ent => new FileEntity(ent.Name, ent.PathFrom)));

            }

            DataContext = this;
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        /// <summary>
        /// Имя текущего конфига
        /// </summary>
        public string SettingName
        {
            get => _settingName;
            set
            {
                if (value != _settingName)
                {
                    _settingName = value;
                    OnPropertyChanged();

                    // Меняю настройки кликабельности кнопок
                    if (!string.IsNullOrWhiteSpace(_settingName) && _settingName.Length > 5)
                        _canRunByName = true;
                    else
                        _canRunByName = false;

                    BtnEnableSwitch();
                }
            }
        }

        /// <summary>
        /// Общий путь для сохранения
        /// </summary>
        public string SharedPathTo
        {
            get => _sharedPathTo;
            set
            {
                if (value != _sharedPathTo)
                {
                    _sharedPathTo = value;
                    OnPropertyChanged();

                    // Меняю настройки кликабельности кнопок
                    if (!string.IsNullOrWhiteSpace(_sharedPathTo) && _sharedPathTo.Length > 10)
                        _canRunByPathTo = true;
                    else
                        _canRunByPathTo = false;

                    BtnEnableSwitch();
                }
            }
        }

        /// <summary>
        /// Ссылка на текущий конфиг
        /// </summary>
        public DBRevitDocExchanges CurrentDBRevitDocExchanges { get; private set; }

        /// <summary>
        /// Ссылка на коллекцию путей к файлам в конфиге
        /// </summary>
        public ObservableCollection<FileEntity> FileEntitiesList { get; private set; } = new ObservableCollection<FileEntity>();

        /// <summary>
        /// Доп. настройки - UserControl
        /// </summary>
        public UserControl SelectedConfig { get; private set; }

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

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

        /// <summary>
        /// Динамическое создание блока доп. параметров (добавление в окно wpf), в зависимости от типа экспорта и уже преднастроенного DBConfigEntity
        /// </summary>
        private void SetExtraSettings_DBConfigEntity(DBConfigEntity dBConfigEntity)
        {
            switch (_revitDocExchangeEnum)
            {
                case (RevitDocExchangeEnum.Navisworks):
                    if (dBConfigEntity is DBNWConfigData dbNWConfigData)
                        SelectedConfig = new NWExtraSettings(dbNWConfigData);
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

        #region Добавление/удаление файлов
        private void OnWindowsAddFile(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Revit Files (*.rvt)|*.rvt",
                InitialDirectory = _initialDirectoryForOpenFileDialog,
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _initialDirectoryForOpenFileDialog = openFileDialog.FileName;
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    FileEntitiesList.Add(new FileEntity(fileName, filePath));
                }
            }
        }

        private void OnBtnHandleAddFile(object sender, RoutedEventArgs e)
        {
            UserStringInput userStringInput = new UserStringInput();
            userStringInput.ShowDialog();
            if (userStringInput.IsRun)
            {
                FileEntitiesList.Add(new FileEntity(userStringInput.UserInputName, userStringInput.UserInputPath));
            }
        }

        private void LBMenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is FileEntity fileEntity)
                    FileEntitiesList.Remove(fileEntity);
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

            // Настройка CurrentDBRevitDocExchanges. Если её нет, то создаём с нуля, иначе - делаем уточнение по параметрам
            if (CurrentDBRevitDocExchanges == null)
            {
                CurrentDBRevitDocExchanges = new DBRevitDocExchanges
                {
                    ProjectId = _project.Id,
                    RevitDocExchangeType = _revitDocExchangeEnum.ToString(),
                    SettingName = this.SettingName,
                    SettingDBFilePath = _sqliteService.CurrentDBFullPath,
                    DescriptionForShow = $"Путь для сохранения: {SharedPathTo}\nКол-во файлов/сборок: {FileEntitiesList.Count}"
                };
            }
            else
            {
                CurrentDBRevitDocExchanges = new DBRevitDocExchanges
                {
                    Id = CurrentDBRevitDocExchanges.Id,
                    ProjectId = _project.Id,
                    RevitDocExchangeType = _revitDocExchangeEnum.ToString(),
                    SettingName = this.SettingName,
                    SettingDBFilePath = _sqliteService.CurrentDBFullPath,
                    DescriptionForShow = $"Путь для сохранения: {SharedPathTo}\nКол-во файлов/сборок: {FileEntitiesList.Count}",
                };
            }

            // Создаю БД для записи Items, а также вношу запись в основную БД диспетчера. Если база ранее была создана - то данный этап игнорирую
            if (!File.Exists(_sqliteService.CurrentDBFullPath))
            {
                _sqliteService.CreateDbFile();
                int idFromDB = _revitDocExchangestDbService.CreateDBRevitDocExchanges(CurrentDBRevitDocExchanges);
                CurrentDBRevitDocExchanges.Id = idFromDB;
            }

            //Создание экземпляра класса на основе введенных данных
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    NWExtraSettings extraSettings = (NWExtraSettings)SelectedConfig;
                    DBNWConfigData sharedDBNWConfigData = extraSettings.CurrentDBNWConfigData;

                    IEnumerable<DBNWConfigData> dBNWConfigDatas = FileEntitiesList
                        .Select(fe => new DBNWConfigData(fe.Name, fe.Path, SharedPathTo).MergeWithDBConfigEntity(sharedDBNWConfigData));
                    
                    if (_sqliteService.GetConfigItems().Count() == 0)
                        _sqliteService.PostConfigItems_ByNWConfigs(dBNWConfigDatas);
                    else
                    {
                        _sqliteService.DropTable();
                        _revitDocExchangestDbService.UpdateDBRevitDocExchanges_ByDBRevitDocExchange(CurrentDBRevitDocExchanges);
                        _sqliteService.PostConfigItems_ByNWConfigs(dBNWConfigDatas);
                    }
                    break;
                case RevitDocExchangeEnum.RevitServer:
                    IEnumerable<DBRSConfigData> dBRSConfigDatas = FileEntitiesList
                        .Select(fe => new DBRSConfigData(fe.Name, fe.Path, SharedPathTo));

                    if (_sqliteService.GetConfigItems().Count() == 0)
                        _sqliteService.PostConfigItems_ByRSConfigs(dBRSConfigDatas);
                    else
                    {
                        _sqliteService.DropTable();
                        _revitDocExchangestDbService.UpdateDBRevitDocExchanges_ByDBRevitDocExchange(CurrentDBRevitDocExchanges);
                        _sqliteService.PostConfigItems_ByRSConfigs(dBRSConfigDatas);
                    }
                    break;
            }

            this.Close();
        }
        #endregion
    }
}
