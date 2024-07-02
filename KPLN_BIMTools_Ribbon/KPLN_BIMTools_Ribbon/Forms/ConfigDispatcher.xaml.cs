using KPLN_BIMTools_Ribbon.Core.SQLite;
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

        private List<CheckBox> _checkedCheckBoxes = new List<CheckBox>();
        private ObservableCollection<DBRevitDocExchanges> _currentDBRevitDocExchanges = new ObservableCollection<DBRevitDocExchanges>();

        public ConfigDispatcher(Logger logger, RevitDocExchangestDbService revitDocExchangestDbService, DBProject project, RevitDocExchangeEnum revitDocExchangeEnum)
        {
            _logger = logger;
            _revitDocExchangestDbService = revitDocExchangestDbService;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;

            InitializeComponent();

            CurrentDBRevitDocExchanges = new ObservableCollection<DBRevitDocExchanges>(_revitDocExchangestDbService.GetDBRevitDocExchanges_ByExchangeTypeANDDBProject(_revitDocExchangeEnum, _project));
            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
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
            UpdateCheckedCheckBoxes(this.iControllGroups);
            BtnEnableSwitch();
        }

        private void OnConfigUnChecked(object sender, RoutedEventArgs e)
        {
            UpdateCheckedCheckBoxes(this.iControllGroups);
            BtnEnableSwitch();
        }

        private void OnBtnRun(object sender, RoutedEventArgs e)
        {
            IsRun = true;
            UpdateCheckedCheckBoxes(this.iControllGroups);
            foreach (CheckBox chkBox in _checkedCheckBoxes)
            {
                if (chkBox.DataContext is DBRevitDocExchanges docExchanges)
                    SelectedDBExchangeEntities.Add(docExchanges);
                else
                    throw new Exception("Скинь разработчику: Не удалось преобразовать тип CheckBox в тип из БД DBRevitDocExchanges");
            }

            this.Close();
        }

        private void OnBtnAddConf(object sender, RoutedEventArgs e)
        {
            // Создание конфига
            FileInfo db_FI = DBEnvironment.GenerateNewPath(_project, _revitDocExchangeEnum);
            SQLiteService sqliteService = new SQLiteService(_logger, db_FI.FullName, _revitDocExchangeEnum);

            ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, _project, _revitDocExchangeEnum);
            configItem.ShowDialog();
            if (configItem.IsRun)
                CurrentDBRevitDocExchanges.Add(configItem.CurrentDBRevitDocExchanges);
        }


        private void MenuItem_Update_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchanges docExchanges)
                {
                    SQLiteService sqliteService = new SQLiteService(_logger, docExchanges.SettingDBFilePath, _revitDocExchangeEnum);
                    ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, _project, _revitDocExchangeEnum, docExchanges);
                    
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

                    ConfigItem configItem = new ConfigItem(_logger, _revitDocExchangestDbService, sqliteService, _project, _revitDocExchangeEnum, docExchanges);
                    configItem.SettingName = $"{configItem.SettingName}_new_{DateTime.Now:d}";

                    configItem.ShowDialog();

                    CurrentDBRevitDocExchanges.Add(configItem.CurrentDBRevitDocExchanges);
                }
            }
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchanges docExchanges)
                    DeleteDBRevitDocExchange(docExchanges);
            }
        }

        private void OnBtnDelConf(object sender, RoutedEventArgs e)
        {
            UserDialog cd = new UserDialog("ВНИМАНИЕ", "Сейчас будут удалены выбранные элементы. Продолжить?");
            cd.ShowDialog();

            if (cd.IsRun)
            {
                List<CheckBox> checkedCheckBoxesToDel = new List<CheckBox>(_checkedCheckBoxes.Count());
                foreach (CheckBox chkBox in _checkedCheckBoxes)
                {
                    if (chkBox.DataContext is DBRevitDocExchanges docExchanges)
                    {
                        DeleteDBRevitDocExchange(docExchanges);
                        checkedCheckBoxesToDel.Add(chkBox);
                    }
                    else
                        throw new Exception("Скинь разработчику: Не удалось преобразовать тип CheckBox в тип из БД DBRevitDocExchanges");
                }

                // Блокирую старт, т.к. ничего не выбрано
                _checkedCheckBoxes.RemoveAll(chBx => checkedCheckBoxesToDel.Contains(chBx));
                BtnEnableSwitch();
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
        /// Обновить коллекцию отмеченных CheckBoxe
        /// </summary>
        private void UpdateCheckedCheckBoxes(ItemsControl itemsControl)
        {
            // Перебор всех элементов в ItemsControl
            for (int i = 0; i < itemsControl.Items.Count; i++)
            {
                // Получение контейнера для элемента
                var container = itemsControl.ItemContainerGenerator.ContainerFromIndex(i) as FrameworkElement;

                // Поиск CheckBox в контейнере
                CheckBox checkBox = FindVisualChild<CheckBox>(container);

                // Проверка и добавление в коллекцию, если CheckBox отмечен
                if (checkBox != null && checkBox.IsChecked == true && !_checkedCheckBoxes.Contains(checkBox))
                {
                    _checkedCheckBoxes.Add(checkBox);
                }

                // Проверка и удаление из коллекции, если CheckBox НЕ отмечен
                if (checkBox != null && checkBox.IsChecked == false && _checkedCheckBoxes.Contains(checkBox))
                {
                    _checkedCheckBoxes.Remove(checkBox);
                }
            }
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
                {
                    return (T)child;
                }
                else
                {
                    T childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                    {
                        return childOfChild;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            if (_checkedCheckBoxes.Count > 0)
            {
                btnRun.IsEnabled = true;
                btnDelConf.IsEnabled = true;
            }
            else
            {
                btnRun.IsEnabled = false;
                btnDelConf.IsEnabled = false;
            }
        }
    }
}
