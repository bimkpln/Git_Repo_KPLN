using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class ConfigDispatcher : Window
    {
        private readonly Logger _logger;
        private readonly RevitDocExchangestDbService _revitDocExchangestDbService;
        private readonly DBProject _project;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;
        private readonly int _revitVersion;

        private readonly List<DBRevitDocExchanges> _selectedDocExchanges = new List<DBRevitDocExchanges>();
        private readonly ObservableCollection<DBRevitDocExchanges> _nativeDBRevitDocExchanges;

        public ConfigDispatcher(
            Logger logger, 
            RevitDocExchangestDbService 
            revitDocExchangestDbService,
            DBProject project, 
            RevitDocExchangeEnum revitDocExchangeEnum,
            int revitVersion)
        {
            _logger = logger;
            _revitDocExchangestDbService = revitDocExchangestDbService;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;
            _revitVersion = revitVersion;

            InitializeComponent();

            _nativeDBRevitDocExchanges = new ObservableCollection<DBRevitDocExchanges>(
                _revitDocExchangestDbService
                .GetDBRevitDocExchanges_ByExchangeTypeANDDBProject(_revitDocExchangeEnum, _project)
                .OrderBy(dExc => dExc.SettingName));
            CurrentDBRevitDocExchanges = new ObservableCollection<DBRevitDocExchanges>(_nativeDBRevitDocExchanges);

            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            FilterItems.Focus();
        }

        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        public bool IsRun { get; private set; }

        /// <summary>
        /// Ссылка на коллекцию конфигов
        /// </summary>
        public ObservableCollection<DBRevitDocExchanges> CurrentDBRevitDocExchanges { get; private set; }

        /// <summary>
        /// Ссылка на выбранные конфиги
        /// </summary>
        public List<DBRevitDocExchanges> SelectedDBExchangeEntities { get; private set; } = new List<DBRevitDocExchanges>();

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void OnConfigChecked(object sender, RoutedEventArgs e)
        {
            UpdateCheckedCheckBoxes();
            BtnEnableSwitch();
        }

        private void OnConfigUnChecked(object sender, RoutedEventArgs e)
        {
            UpdateCheckedCheckBoxes();
            BtnEnableSwitch();
        }

        private void OnBtnRun(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            
            UpdateCheckedCheckBoxes();
            
            foreach (DBRevitDocExchanges docExch in _selectedDocExchanges)
            {
                SelectedDBExchangeEntities.Add(docExch);
            }

            this.Close();
        }

        private void OnBtnAddConf(object sender, RoutedEventArgs e)
        {
            // Создание конфига
            FileInfo db_FI = DBEnvironment.GenerateNewPath(_project, _revitDocExchangeEnum);
            SQLiteService sqliteService = new SQLiteService(_logger, db_FI.FullName, _revitDocExchangeEnum);

            ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, _project, _revitDocExchangeEnum, _revitVersion);
            if ((bool)configItem.ShowDialog())
            {
                CurrentDBRevitDocExchanges.Add(configItem.CurrentDBRevitDocExchanges);
                SortCurrentDBRevitDocExchanges();
            }
        }

        /// <summary>
        /// Обновить сортировку можно только перезаписью коллекции. Сортировка не даёт сигнал на обновление
        /// </summary>
        private void SortCurrentDBRevitDocExchanges()
        {
            var sortedList = CurrentDBRevitDocExchanges.OrderBy(dExc => dExc.SettingName).ToList();
            CurrentDBRevitDocExchanges.Clear();
            foreach (var item in sortedList)
            {
                CurrentDBRevitDocExchanges.Add(item);
            }
        }

        private void OnBtnDelConf(object sender, RoutedEventArgs e)
        {
            UserDialog cd = new UserDialog("ВНИМАНИЕ", "Сейчас будут удалены выбранные элементы. Продолжить?");
            cd.ShowDialog();

            if (cd.IsRun)
            {
                foreach (DBRevitDocExchanges docExch in _selectedDocExchanges)
                {
                    DeleteDBRevitDocExchange(docExch);
                }

                // Блокирую старт, т.к. ничего не выбрано
                _selectedDocExchanges.Clear();
                BtnEnableSwitch();
            }
        }

        private void MenuItem_Update_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchanges docExchanges)
                {
                    SQLiteService sqliteService = new SQLiteService(_logger, docExchanges.SettingDBFilePath, _revitDocExchangeEnum);
                    ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, _project, 
                        _revitDocExchangeEnum, _revitVersion, docExchanges);
                    
                    configItem.ShowDialog();

                    // Обновляю основную коллекцию новыми данными
                    int index = CurrentDBRevitDocExchanges.IndexOf(docExchanges);
                    if (index >= 0)
                    {
                        // Уведомить об изменении элемента
                        CurrentDBRevitDocExchanges[index] = configItem.CurrentDBRevitDocExchanges;
                    }
                }
            }
        }

        private void MenuItem_Copy_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchanges docExchanges)
                {
                    // Создание конфига
                    FileInfo db_FI = DBEnvironment.GenerateNewPath(_project, _revitDocExchangeEnum);
                    SQLiteService sqliteService = new SQLiteService(_logger, db_FI.FullName, _revitDocExchangeEnum);

                    ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, 
                        _project, _revitDocExchangeEnum, _revitVersion, docExchanges);
                    configItem.SettingName = $"{configItem.SettingName}_new_{DateTime.Now:d}";

                    bool? dialogResult = configItem.ShowDialog();
                    if ((bool)dialogResult)
                        CurrentDBRevitDocExchanges.Add(configItem.CurrentDBRevitDocExchanges);
                }
            }
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchanges docExchanges)
                {
                    if (_selectedDocExchanges.Count > 1)
                    {
                        UserDialog cd = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалено {_selectedDocExchanges.Count} конфигурации. Продолжить?");
                        cd.ShowDialog();
                        if (cd.IsRun)
                        {
                            foreach (DBRevitDocExchanges docEcxh in _selectedDocExchanges)
                            {
                                DeleteDBRevitDocExchange(docEcxh);
                            }

                            // Блокирую старт, т.к. ничего не выбрано
                            _selectedDocExchanges.Clear();
                            BtnEnableSwitch();
                        }
                    }
                    else
                    {
                        UserDialog cd = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалена конфигурация \"{docExchanges.SettingName}\". Продолжить?");
                        cd.ShowDialog();
                    
                        if (cd.IsRun)
                            DeleteDBRevitDocExchange(docExchanges);
                    }
                }
            }
        }

        private void DeleteDBRevitDocExchange(DBRevitDocExchanges docExchanges)
        {
            SelectedDBExchangeEntities.Remove(docExchanges);

            _revitDocExchangestDbService.DeleteDBRevitDocExchange_ById(docExchanges.Id);
            CurrentDBRevitDocExchanges.Remove(docExchanges);

            FileInfo currentDB = new FileInfo(docExchanges.SettingDBFilePath);
            currentDB.Delete();
        }

        /// <summary>
        /// Обновить коллекцию отмеченных CheckBoxes
        /// </summary>
        /// <param name="isReloadDataContext">Флаг обновления основной коллекции</param>
        private void UpdateCheckedCheckBoxes(bool isReloadDataContext = false)
        {
            // InvokeAsync - чтобы дать системе время на завершение отрисовки визуальных элементов перед тем, как пытаться искать чекбоксы
            Dispatcher.InvokeAsync(() =>
            {
                // Перебор всех элементов в ItemsControl
                for (int i = 0; i < this.iControllGroups.Items.Count; i++)
                {
                    // Получение контейнера для элемента
                    if (this.iControllGroups.ItemContainerGenerator.ContainerFromIndex(i) is FrameworkElement container)
                    {
                        CheckBox checkBox = FindVisualChild<CheckBox>(container);
                        if (checkBox != null && checkBox.DataContext is DBRevitDocExchanges docExchange)
                        {
                            // Если основная коллекция обновляется, нужно брать данные по ранее добавленым элементам из кэша (CheckBoxы все заново строятся)
                            if (isReloadDataContext)
                                checkBox.IsChecked = _selectedDocExchanges.Any(docExch => docExch.Id == docExchange.Id);
                            else if (checkBox.IsChecked == true && !_selectedDocExchanges.Any(docExch => docExch.Id == docExchange.Id))
                                _selectedDocExchanges.Add(docExchange);
                            else if (!isReloadDataContext && checkBox.IsChecked == false)
                                _selectedDocExchanges.Remove(docExchange);
                        }
                    }
                }
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// Поиск зависимых объектов 
        /// </summary>
        private T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }

            return null;
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            // InvokeAsync - чтобы дать системе время на завершение отрисовки визуальных элементов перед тем, как пытаться искать чекбоксы
            Dispatcher.InvokeAsync(() =>
            {
                if (_selectedDocExchanges.Count > 0)
                {
                    btnRun.IsEnabled = true;
                    btnDelConf.IsEnabled = true;
                }
                else
                {
                    btnRun.IsEnabled = false;
                    btnDelConf.IsEnabled = false;
                }
            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        private void FilterItems_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox tBox = sender as TextBox; 
            string userInput = tBox.Text.ToLower();
            CurrentDBRevitDocExchanges.Clear();

            DBRevitDocExchanges[] docExchangesNotContaines = _nativeDBRevitDocExchanges.Where(de =>  !de.SettingName.ToLower().Contains(userInput)).ToArray();
            foreach(DBRevitDocExchanges docExcOut in docExchangesNotContaines)
            {
                CurrentDBRevitDocExchanges.Remove(docExcOut);
            }

            DBRevitDocExchanges[] docExchangesContaines = _nativeDBRevitDocExchanges.Where(de => de.SettingName.ToLower().Contains(userInput)).ToArray();
            foreach (DBRevitDocExchanges docExcIn in docExchangesContaines)
            {
                CurrentDBRevitDocExchanges.Add(docExcIn);
            }

            UpdateCheckedCheckBoxes(true);
        }
    }
}
