using KPLN_BIMTools_Ribbon.Common;
using KPLN_BIMTools_Ribbon.Core.SQLite;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_Forms.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class ConfigDispatcher : Window
    {
        private readonly DBModuleAutostart[] _dbModuleAutostarArrForUser;
        private readonly ObservableCollection<DBRevitDocExchangesWrapper> _dbRevitDocExchWrappers;

        private readonly Logger _logger;
        private readonly DBProject _project;
        private readonly RevitDocExchangeEnum _revitDocExchangeEnum;
        /// <summary>
        /// Метка режима запуска - создание конфигурации для автозапуска
        /// </summary>
        private readonly bool _isAutoStartConfig;
        private readonly int _moduleId = 80;

        public ConfigDispatcher(
            Logger logger,
            DBProject project,
            RevitDocExchangeEnum revitDocExchangeEnum,
            bool isAutoStartConfig)
        {
            _logger = logger;
            _project = project;
            _revitDocExchangeEnum = revitDocExchangeEnum;
            _isAutoStartConfig = isAutoStartConfig;

            InitializeComponent();


            // Получаю исходную версию сущностей из БД
            _dbRevitDocExchWrappers = new ObservableCollection<DBRevitDocExchangesWrapper>(ExchangeService.CurrentRevitDocExchangesDbService
                    .GetDBRevitDocExchanges_ByExchangeTypeANDDBProject(_revitDocExchangeEnum, _project)
                    .OrderBy(dExc => dExc.SettingName)
                    .Select(dExc => new DBRevitDocExchangesWrapper(dExc)));

            FiltereDBRevitDocExchdWrappers = CollectionViewSource.GetDefaultView(_dbRevitDocExchWrappers);
            FiltereDBRevitDocExchdWrappers.Filter = FilterItemsPredicate;

            // Устанавливаю описание/коллекции в зависимости от алгоритма запуска
            if (_isAutoStartConfig)
            {
                _dbModuleAutostarArrForUser = ExchangeService
                    .CurrentModuleAutostartDbService
                    .GetDBModuleAutostartsByUserAndRVersionAndPrjIdAndTable(ExchangeService.CurrentDBUser.Id, Module.RevitVersion, _project.Id, _moduleId, DB_Enumerator.RevitDocExchanges.ToString())
                    .ToArray();

                // Взвожу галку, если конфиг в списке
                foreach(DBModuleAutostart dBModuleAutostart in _dbModuleAutostarArrForUser)
                {
                    DBRevitDocExchangesWrapper selectedExchWr = _dbRevitDocExchWrappers.FirstOrDefault(dExhWr => dExhWr.Id == dBModuleAutostart.DBTableKeyId);
                    if (selectedExchWr == null) continue;

                    selectedExchWr.IsSelected = true;
                }


                btnRun.ToolTip = "Сохранить выбранные конфигурации в список на автозапуск";
                btnRun.Content = "🖬";
            }
            else
            {
                btnRun.ToolTip = "Запустить выбранные конфигурации";
                btnRun.Content = "▶";
            }

            DataContext = this;

            PreviewKeyDown += new KeyEventHandler(HandleEsc);
            FilterItems.Focus();
            BtnEnableSwitch();
        }

        public ICollectionView FiltereDBRevitDocExchdWrappers { get; }

        /// <summary>
        /// Ссылка на выбранные конфиги
        /// </summary>
        public DBRevitDocExchangesWrapper[] SelectedDBExchWrappers => _dbRevitDocExchWrappers.Where(ent => ent.IsSelected).ToArray();

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void OnConfigClicked(object sender, RoutedEventArgs e) => BtnEnableSwitch();

        private void OnBtnRun(object sender, RoutedEventArgs e)
        {
            if (_isAutoStartConfig)
            {
                // Удаляю НЕ отмеченные (их сняли)
                foreach (DBModuleAutostart dBModuleAutostart in _dbModuleAutostarArrForUser)
                {
                    DBRevitDocExchangesWrapper selectedExchWr = _dbRevitDocExchWrappers.FirstOrDefault(dExhWr => dExhWr.Id == dBModuleAutostart.DBTableKeyId);
                    if (selectedExchWr == null) continue;

                    if (!selectedExchWr.IsSelected)
                        ExchangeService
                            .CurrentModuleAutostartDbService
                            .DeleteDBModuleAutostarts(dBModuleAutostart);
                }


                // Создаю и обновляю новые
                var selectedDocExch = SelectedDBExchWrappers
                    .Select(docExch => new DBModuleAutostart()
                        {
                            UserId = ExchangeService.CurrentDBUser.Id,
                            RevitVersion = Module.RevitVersion,
                            ProjectId = _project.Id,
                            ModuleId = _moduleId,
                            DBTableName = DB_Enumerator.RevitDocExchanges.ToString(),
                            DBTableKeyId = docExch.Id,
                        });

                ExchangeService
                    .CurrentModuleAutostartDbService
                    .BulkCreateDBModuleAutostarts(selectedDocExch);
            }

            DialogResult = true;
            this.Close();
        }

        private void OnBtnAddConf(object sender, RoutedEventArgs e)
        {
            // Создание конфига
            FileInfo db_FI = DBEnvironment.GenerateNewPath(_project, _revitDocExchangeEnum);
            SQLiteService sqliteService = new SQLiteService(_logger, db_FI.FullName, _revitDocExchangeEnum);

            ConfigItem configItem = new ConfigItem(_logger, sqliteService, _project, _revitDocExchangeEnum);
            if ((bool)configItem.ShowDialog() && configItem.DBRevitDocExchWrapper is DBRevitDocExchangesWrapper dExchEnt)
            {
                _dbRevitDocExchWrappers.Add(dExchEnt);
                SortCurrentDBRevitDocExchanges();
            }
        }

        /// <summary>
        /// Обновить сортировку можно только перезаписью коллекции. Сортировка не даёт сигнал на обновление
        /// </summary>
        private void SortCurrentDBRevitDocExchanges()
        {
            var sortedList = _dbRevitDocExchWrappers.OrderBy(dExc => dExc.SettingName).ToList();
            _dbRevitDocExchWrappers.Clear();
            foreach (var item in sortedList)
            {
                _dbRevitDocExchWrappers.Add(item);
            }
        }

        private void OnBtnDelConf(object sender, RoutedEventArgs e)
        {
            UserDialog cd = new UserDialog("ВНИМАНИЕ", "Сейчас будут удалены выбранные элементы. Продолжить?");
            cd.ShowDialog();

            if (cd.IsRun)
            {
                DeleteDBRevitDocExchange(SelectedDBExchWrappers);

                BtnEnableSwitch();
            }
        }

        private void MenuItem_Update_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchangesWrapper docExchangeEnt)
                {
                    SQLiteService sqliteService = new SQLiteService(_logger, docExchangeEnt.SettingDBFilePath, _revitDocExchangeEnum);
                    ConfigItem configItem = new ConfigItem(_logger, sqliteService, _project,
                        _revitDocExchangeEnum, docExchangeEnt);

                    configItem.ShowDialog();

                    // Обновляю основную коллекцию новыми данными
                    int index = _dbRevitDocExchWrappers.IndexOf(docExchangeEnt);
                    if (index >= 0)
                        // Уведомить об изменении элемента
                        _dbRevitDocExchWrappers[index] = configItem.DBRevitDocExchWrapper;
                }
            }
        }

        private void MenuItem_Copy_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchangesWrapper docExchangeEnt)
                {
                    // Создание конфига
                    FileInfo db_FI = DBEnvironment.GenerateNewPath(_project, _revitDocExchangeEnum);
                    SQLiteService sqliteService = new SQLiteService(_logger, db_FI.FullName, _revitDocExchangeEnum);

                    ConfigItem configItem = new ConfigItem(_logger, sqliteService,
                        _project, _revitDocExchangeEnum, docExchangeEnt);
                    configItem.SettingName = $"{configItem.SettingName}_new_{DateTime.Now:d}";

                    bool? dialogResult = configItem.ShowDialog();
                    if ((bool)dialogResult)
                    {
                        _dbRevitDocExchWrappers.Add(configItem.DBRevitDocExchWrapper);
                        SortCurrentDBRevitDocExchanges();
                    }
                }
            }
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            if ((MenuItem)e.Source is MenuItem menuItem)
            {
                if (menuItem.DataContext is DBRevitDocExchangesWrapper docExchWrapper)
                {
                    if (SelectedDBExchWrappers.Count() > 1)
                    {
                        UserDialog cd = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалено {SelectedDBExchWrappers.Count()} конфигурации. Продолжить?");
                        cd.ShowDialog();
                        if (cd.IsRun)
                        {
                            DeleteDBRevitDocExchange(SelectedDBExchWrappers);
                            foreach (DBRevitDocExchangesWrapper docEcxhWr in SelectedDBExchWrappers)

                            // Блокирую старт, т.к. ничего не выбрано
                            BtnEnableSwitch();
                        }
                    }
                    else
                    {
                        UserDialog cd = new UserDialog("ВНИМАНИЕ", $"Сейчас будут удалена конфигурация \"{docExchWrapper.SettingName}\". Продолжить?");
                        cd.ShowDialog();

                        if (cd.IsRun)
                            DeleteDBRevitDocExchange(new DBRevitDocExchangesWrapper[] { docExchWrapper });
                    }
                }
            }
        }

        private void DeleteDBRevitDocExchange(IEnumerable<DBRevitDocExchangesWrapper> docExchWrappers)
        {
            ExchangeService.CurrentRevitDocExchangesDbService.DeleteDBRevitDocExchange_ByIdColl(docExchWrappers.Select(docExch => docExch.Id));
            
            foreach(DBRevitDocExchangesWrapper docExchWrapper in docExchWrappers)
            {
                _dbRevitDocExchWrappers.Remove(docExchWrapper);

                FileInfo currentDB = new FileInfo(docExchWrapper.SettingDBFilePath);
                currentDB.Delete();
            }
        }

        /// <summary>
        /// Свитчер кликабельности основных кнопок управления
        /// </summary>
        private void BtnEnableSwitch()
        {
            // InvokeAsync - чтобы дать системе время на завершение отрисовки визуальных элементов перед тем, как пытаться искать чекбоксы
            Dispatcher.InvokeAsync(() =>
            {
                if (SelectedDBExchWrappers.Count() > 0)
                {
                    btnRun.IsEnabled = true;
                    btnDelConf.IsEnabled = true;
                }
                else
                {
                    btnRun.IsEnabled = false;
                    btnDelConf.IsEnabled = false;
                }

                if (_isAutoStartConfig)
                    btnRun.IsEnabled = true;

            }, System.Windows.Threading.DispatcherPriority.Background);
        }

        /// <summary>
        /// Метод фильтрации коллекции в окне
        /// </summary>
        private bool FilterItemsPredicate(object item)
        {
            if (item is DBRevitDocExchangesWrapper docExchWrapper)
            {
                bool isOnlyChecked = (bool)this.OnlySelChBx.IsChecked;
                
                string userStringInput = this.FilterItems.Text.ToLower();
                if (string.IsNullOrEmpty(userStringInput) || docExchWrapper.SettingName.ToLower().Contains(userStringInput))
                {
                    if (isOnlyChecked)
                        return docExchWrapper.IsSelected;

                    return true;
                }
            }

            return false;
        }

        private void FilterItems_TextChanged(object sender, TextChangedEventArgs e) => FiltereDBRevitDocExchdWrappers?.Refresh();

        private void OnlySelChBx_Checked(object sender, RoutedEventArgs e) => FiltereDBRevitDocExchdWrappers?.Refresh();
    }
}
