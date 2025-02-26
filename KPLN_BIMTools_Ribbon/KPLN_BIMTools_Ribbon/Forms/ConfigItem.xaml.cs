using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Core.SQLite.Entities;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using Microsoft.Win32;
using NLog;
using RevitServerAPILib;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
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

        private readonly NLog.Logger _logger;
        private readonly RevitDocExchangestDbService _revitDocExchangestDbService;
        private readonly SQLiteService _sqliteService;
        private readonly DBProject _project;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;
        private readonly int _revitVersion;

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
            NLog.Logger logger,
            RevitDocExchangestDbService revitDocExchangestDbService,
            SQLiteService sqliteService,
            DBProject project,
            RevitDocExchangeEnum revitDocExchangeEnum,
            int revitVersion,
            DBRevitDocExchanges currentDBRevitDocExchanges = null)
        {
            _logger = logger;
            _revitDocExchangestDbService = revitDocExchangestDbService;
            _sqliteService = sqliteService;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;
            _revitVersion = revitVersion;
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
                if (dBConfigEntities.Any())
                    SetExtraSettings_DBConfigEntity(dBConfigEntities.FirstOrDefault());
                else
                    SetExtraSettings();

                SharedPathTo = CurrentDBRevitDocExchanges.SettingResultPath;

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

                    if(Directory.Exists(_sharedPathTo))
                        _initialDirectoryForOpenFileDialog = _sharedPathTo;
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
        public UserControl SelectedConfig 
        { 
            get; 
            private set; 
        }

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
                    SelectedConfig = new RVTExtraSettings(new DBRVTConfigData()
                    {
                        MaxBackup = 10,
                    });
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
                    if (dBConfigEntity is DBRVTConfigData dbRVTConfigData)
                        SelectedConfig = new RVTExtraSettings(dbRVTConfigData);
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

        /// <summary>
        /// Обработка прокрутки колесика на части окна с файлами
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FileWrapPanel_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Передаём событие ScrollViewer'у
            if (FileStackScroll != null)
            {
                // Если крутим вниз, увеличиваем вертикальное смещение
                if (e.Delta < 0)
                {
                    FileStackScroll.ScrollToVerticalOffset(FileStackScroll.VerticalOffset + 20); // Примерный шаг прокрутки
                }
                // Если крутим вверх, уменьшаем вертикальное смещение
                else if (e.Delta > 0)
                {
                    FileStackScroll.ScrollToVerticalOffset(FileStackScroll.VerticalOffset - 20);
                }

                // Указываем, что событие обработано
                e.Handled = true;
            }
        }

        private void OnMainPathAddFolder(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.FolderBrowserDialog openFolderDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                string oldSelectedPath = this.PathTo.Text;
                if (string.IsNullOrEmpty(oldSelectedPath))
                    openFolderDialog.SelectedPath = _initialDirectoryForOpenFileDialog;
                else
                    openFolderDialog.SelectedPath = oldSelectedPath;
                
                if (openFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK 
                    && !string.IsNullOrWhiteSpace(openFolderDialog.SelectedPath))
                {
                    SharedPathTo = openFolderDialog.SelectedPath;
                }
            }
        }

        private void OnMainPathAddRevitServerFolder(object sender, RoutedEventArgs e)
        {
            ElementSinglePick selectedRevitServerMainDirForm = SelectRevitServerMainDir.CreateForm_SelectRSMainDir(_revitVersion);
            bool? dialogResult = selectedRevitServerMainDirForm.ShowDialog();
            if (dialogResult == null || selectedRevitServerMainDirForm.Status != UIStatus.RunStatus.Run)
                return;

            string selectedRSMainDirFullPath = selectedRevitServerMainDirForm.SelectedElement.Element as string;
            string selectedRSHostName = selectedRSMainDirFullPath.Split('\\')[0];
            string selectedRSMainDir = selectedRSMainDirFullPath.TrimStart(selectedRSHostName.ToCharArray());

            RevitServer revitServer = new RevitServer(selectedRSHostName, _revitVersion);

            IList<Folder> rsFolders = revitServer.GetFolderContents(selectedRSMainDir, 0).Folders;
            List<ElementEntity> activeEntitiesForForm = new List<ElementEntity>(
                rsFolders
                    .Where(f => f.LockState != LockState.Locked)
                    .Select(f => new ElementEntity(f.Path))
                    .ToArray());

            ElementSinglePick pickForm = new ElementSinglePick(activeEntitiesForForm.OrderBy(p => p.Name), "Выбери папку Revit-Server");
            bool? pickFormResult = pickForm.ShowDialog();
            if (pickFormResult == null || pickForm.Status != UIStatus.RunStatus.Run)
                return;
            
            SharedPathTo = $"\\\\{selectedRSHostName}{pickForm.SelectedElement.Name}";
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
                    AddToFileEntitiesWithCheck(new FileEntity(fileName, filePath));
                }
            }
        }

        private void OnAddNewRevitServerLink(object sender, RoutedEventArgs e)
        {
            ElementMultiPick rsFilesPickForm = SelectFilesFromRevitServer.CreateForm(_revitVersion);
            if (rsFilesPickForm == null)
                return;

            bool? dialogResult = rsFilesPickForm.ShowDialog();
            if (dialogResult == null || rsFilesPickForm.Status != UIStatus.RunStatus.Run)
                return;

            foreach (ElementEntity formEntity in rsFilesPickForm.SelectedElements)
            {
                AddToFileEntitiesWithCheck(new FileEntity(formEntity.Name, $"RSN:\\\\{SelectFilesFromRevitServer.CurrentRevitServer.Host}{formEntity.Name}"));
            }
        }

        private void OnBtnHandleAddFile(object sender, RoutedEventArgs e)
        {
            UserStringInfo userStringInput = new UserStringInfo(false);
            bool? dialogResult = userStringInput.ShowDialog();
            if (dialogResult == null || dialogResult == false)
                return;

            AddToFileEntitiesWithCheck(new FileEntity(userStringInput.UserInputName, userStringInput.UserInputPath));
        }

        private void LBMenuItem_Info_Click(object sender, RoutedEventArgs e)
        {
            // Получаем выбранные элементы
            var selectedItems = fileWrapPanel.SelectedItems.Cast<FileEntity>().ToList();
            if (selectedItems.Count != 1)
            {
                UserDialog cd = new UserDialog("Предупреждение", $"Информацию можно смотерть только по отдельным файлам");
                cd.ShowDialog();
            }
            else
            {
                if (selectedItems.FirstOrDefault() is FileEntity fileEntity)
                {
                    UserStringInfo userStringInput = new UserStringInfo(true, fileEntity.Name, fileEntity.Path);
                    userStringInput.ShowDialog();
                }
            }
        }

        private void LBMenuItem_Edit_Click(object sender, RoutedEventArgs e)
        {
            // Получаем выбранные элементы
            var selectedItems = fileWrapPanel.SelectedItems.Cast<FileEntity>().ToList();
            if (selectedItems.Count != 1)
            {
                UserDialog cd = new UserDialog("Предупреждение", $"Редактировать можно только по отдельным файлам");
                cd.ShowDialog();
            }
            else
            {
                FileEntity selectedItem = selectedItems.FirstOrDefault();
                if (selectedItem is FileEntity fileEntity)
                {
                    UserStringInfo userStringInput = new UserStringInfo(false, fileEntity.Name, fileEntity.Path);
                    userStringInput.ShowDialog();
                    if ((bool)userStringInput.DialogResult)
                    {
                        int updateIndex = FileEntitiesList.IndexOf(fileEntity);
                        FileEntitiesList.Remove(fileEntity);

                        fileEntity.Name = userStringInput.UserInputName;
                        fileEntity.Path = userStringInput.UserInputPath;
                        FileEntitiesList.Insert(updateIndex, fileEntity);
                    }
                }
            }
        }

        private void LBMenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            // Получаем выбранные элементы
            var selectedItems = fileWrapPanel.SelectedItems.Cast<FileEntity>().ToList();

            // Удаляем каждый выбранный элемент из коллекции
            foreach (var item in selectedItems)
            {
                if (item is FileEntity fileEntity)
                    FileEntitiesList.Remove(fileEntity);
            }
        }
        #endregion

        /// <summary>
        /// Добавить эл-т в коллекцию FileEntitiesList с предпроверкой
        /// </summary>
        private void AddToFileEntitiesWithCheck(FileEntity entity)
        {
            if (!FileEntitiesList.Any(listEnt => listEnt.Path.Equals(entity.Path) || listEnt.Name.Equals(entity.Name)))
                FileEntitiesList.Add(entity);
            else
            {
                UserDialog cd = new UserDialog("Предупреждение", $"Файл по указнному пути или с тем же именем уже есть в списке. В добавлении отказано");
                cd.ShowDialog();
            }
        }

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
                    SettingName = SettingName,
                    SettingResultPath = SharedPathTo,
                    SettingCountItem = FileEntitiesList.Count,
                    SettingDBFilePath = _sqliteService.CurrentDBFullPath,
                };
            }
            else
            {
                CurrentDBRevitDocExchanges = new DBRevitDocExchanges
                {
                    Id = CurrentDBRevitDocExchanges.Id,
                    ProjectId = _project.Id,
                    RevitDocExchangeType = _revitDocExchangeEnum.ToString(),
                    SettingName = SettingName,
                    SettingResultPath = SharedPathTo,
                    SettingCountItem = FileEntitiesList.Count,
                    SettingDBFilePath = _sqliteService.CurrentDBFullPath,
                };
            }

            // Создаю БД для записи Items, а также вношу запись в основную БД диспетчера. Если база ранее была создана - то данный этап игнорирую
            if (!System.IO.File.Exists(_sqliteService.CurrentDBFullPath))
            {
                _sqliteService.CreateDbFile();
                int idFromDB = _revitDocExchangestDbService.CreateDBRevitDocExchanges(CurrentDBRevitDocExchanges);
                CurrentDBRevitDocExchanges.Id = idFromDB;
            }

            //Создание экземпляра класса на основе введенных данных
            switch (_revitDocExchangeEnum)
            {
                case RevitDocExchangeEnum.Navisworks:
                    NWExtraSettings extraNWSettings = (NWExtraSettings)SelectedConfig;
                    DBNWConfigData dbNWConfigData = extraNWSettings.CurrentDBNWConfigData;

                    IEnumerable<DBNWConfigData> dBNWConfigDatas = FileEntitiesList
                        .Select(fe => new DBNWConfigData(fe.Name, fe.Path, SharedPathTo).MergeWithDBConfigEntity(dbNWConfigData));

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
                    RVTExtraSettings extraRVTSettings = (RVTExtraSettings)SelectedConfig;
                    DBRVTConfigData dbRVTConfigData = extraRVTSettings.CurrentDBRSConfigData;

                    IEnumerable<DBRVTConfigData> dBRVTConfigDatas = FileEntitiesList
                        .Select(fe => new DBRVTConfigData(fe.Name, fe.Path, SharedPathTo).MergeWithDBConfigEntity(dbRVTConfigData));

                    if (_sqliteService.GetConfigItems().Count() == 0)
                        _sqliteService.PostConfigItems_ByRSConfigs(dBRVTConfigDatas);
                    else
                    {
                        _sqliteService.DropTable();
                        _revitDocExchangestDbService.UpdateDBRevitDocExchanges_ByDBRevitDocExchange(CurrentDBRevitDocExchanges);
                        _sqliteService.PostConfigItems_ByRSConfigs(dBRVTConfigDatas);
                    }
                    break;
            }

            this.Close();
        }
    }
}
