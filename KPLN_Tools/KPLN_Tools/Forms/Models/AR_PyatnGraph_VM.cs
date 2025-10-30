using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Tools.Common.AR_PyatnGraph;
using KPLN_Tools.ExecutableCommand;
using KPLN_Tools.Forms.Models.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;

namespace KPLN_Tools.Forms.Models
{
    /// <summary>
    /// VM для окна ввода данных по пятнографии
    /// </summary>
    public class AR_PyatnGraph_VM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly Document _doc;
        private readonly ObservableCollection<ARPG_TZ_FlatData> _collection = new ObservableCollection<ARPG_TZ_FlatData>();
        private readonly HashSet<ARPG_TZ_FlatData> _subscribedItems = new HashSet<ARPG_TZ_FlatData>();
        private readonly string _cofigName = "AR_PyatnGraph";
        private readonly ConfigType _configType = ConfigType.Shared;

        private double _sumPercent;
        private bool _canRun = false;
        private Brush _sumPercentColor = Brushes.White;

        [JsonConstructor]
        public AR_PyatnGraph_VM() { }


        public AR_PyatnGraph_VM(UIApplication uiapp)
        {
            _doc = uiapp.ActiveUIDocument.Document;

            #region Заполняю поля окна в зависимости от наличия файла конфига
            // Файл конфига присутсвует
            if (ConfigService.ReadConfigFile<ARPG_TZ_Config>(ModuleData.RevitVersion, _doc, _configType, _cofigName) is ARPG_TZ_Config config)
            {
                _collection = config.Config_ARPG_TZ_FlatDataList;
                ARPG_TZ_MainData = config.Config_ARPG_TZ_MainData;
            }
            else
                ARPG_TZ_MainData = new ARPG_TZ_MainData();

            ARPG_TZ_FlatDataColl = CollectionViewSource.GetDefaultView(_collection);
            #endregion

            PresetFlatCodeCommand = new RelayCommand<object>(_ => PresetFlatCode());
            SaveConfigCommand = new RelayCommand<object>(_ => SaveConfig());
            RunCommand = new RelayCommand<object>(_ => Run());
            AddNewFlatDataCommand = new RelayCommand<object>(_ => AddNewFlatData());
            SortByRangeNameCommand = new RelayCommand<object>(_ => SortByText(nameof(ARPG_TZ_FlatData.TZRangeName)));
            SortByFlatCodeCommand = new RelayCommand<object>(_ => SortByText(nameof(ARPG_TZ_FlatData.TZCode)));
            DeleteFlatDataCommand = new RelayCommand<ARPG_TZ_FlatData>(DeleteFlatData);

            // 1) Подписаться на уже существующие элементы (если они есть)
            foreach (var item in _collection)
            {
                SubscribeToItem(item);
            }

            // 2) Подписаться на изменения самой коллекции
            _collection.CollectionChanged += OnCollectionChanged;

            UpdateSumPercent();
        }

        /// <summary>
        /// Ссылка на класс конфига
        /// </summary>
        public ARPG_TZ_Config ARPG_TZ_Config { get; set; }

        /// <summary>
        /// Ссылка на ARPG_MainData_TZ в окне
        /// </summary>
        public ARPG_TZ_MainData ARPG_TZ_MainData { get; set; }

        /// <summary>
        /// Ссылка на класс конфига
        /// </summary>
        public ICollectionView ARPG_TZ_FlatDataColl { get; set; }

        /// <summary>
        /// Сумма процентов в окне
        /// </summary>
        public double SumPercent
        {
            get => _sumPercent;
            set
            {
                if (_sumPercent != value)
                {
                    _sumPercent = value;

                    // Сравниваем с допуском, чтобы избежать проблем с floating point
                    if (Math.Abs(_sumPercent - 100.0) <= 0.0001)
                    {
                        SumPercentColor = Brushes.Green;
                        CanRun = true;
                    }
                    else
                    {
                        SumPercentColor = Brushes.Red;
                        CanRun = false;
                    }

                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Метка возможности запуска
        /// </summary>
        public bool CanRun
        {
            get => _canRun;
            set
            {
                if (_canRun != value)
                {
                    _canRun = value;
                    NotifyPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Цвет суммы процентов в окне
        /// </summary>
        public Brush SumPercentColor
        {
            get => _sumPercentColor;
            set
            {
                if (_sumPercentColor != value)
                {
                    _sumPercentColor = value;
                    NotifyPropertyChanged();
                }
            }
        }

        #region Набор команд
        /// <summary>
        /// Комманда: Сортировать по диапазону
        /// </summary>
        public ICommand SortByRangeNameCommand { get; }

        /// <summary>
        /// Комманда: Сортировать по коду квартиры
        /// </summary>
        public ICommand SortByFlatCodeCommand { get; }

        /// <summary>
        /// Комманда: Предварительное назначение кода квартиры
        /// </summary>
        public ICommand PresetFlatCodeCommand { get; }

        /// <summary>
        /// Комманда: Добавить новый тип квартир
        /// </summary>
        public ICommand AddNewFlatDataCommand { get; }

        /// <summary>
        /// Комманда: Сохранить конфиг
        /// </summary>
        public ICommand SaveConfigCommand { get; }

        /// <summary>
        /// Комманда: Запуск рассчёта
        /// </summary>
        public ICommand RunCommand { get; }

        /// <summary>
        /// Комманда: Удалить тип квартир
        /// </summary>
        public ICommand DeleteFlatDataCommand { get; }
        #endregion

        #region Реализация команд
        /// <summary>
        /// Реализация: Предварительное назначение кода квартиры
        /// </summary>
        private void PresetFlatCode()
        {
            try
            {
                ARPG_Room[] arpgRooms = ARPG_RoomsFromDoc(true);
                ARPG_Flat[] aRPGFlats = ARPG_Flat.Get_ARPG_Flats(ARPG_TZ_MainData, arpgRooms);
                
                if(arpgRooms != null && aRPGFlats != null)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdARPG_PreSetData(_doc, ARPG_TZ_MainData, arpgRooms, aRPGFlats, _collection.ToArray()));
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
            }
        }

        /// <summary>
        /// Реализация: Сохранить конфиг
        /// </summary>
        private void SaveConfig()
        {
            ARPG_TZ_Config = new ARPG_TZ_Config
            {
                Config_ARPG_TZ_FlatDataList = _collection,
                Config_ARPG_TZ_MainData = ARPG_TZ_MainData
            };

            try
            {
                ConfigService.SaveConfig<ARPG_TZ_Config>(ModuleData.RevitVersion, _doc, _configType, ARPG_TZ_Config, _cofigName);

                MessageBox.Show(
                        $"Конфиг успешно сохранён",
                        "Результат",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
            }
        }

        /// <summary>
        /// Реализация: Запуск рассчёта
        /// </summary>
        private void Run()
        {
            try
            {
                ARPG_Room[] arpgRooms = ARPG_RoomsFromDoc(false);
                ARPG_Flat[] aRPGFlats = ARPG_Flat.Get_ARPG_Flats(ARPG_TZ_MainData, arpgRooms);

                if (arpgRooms != null && aRPGFlats != null)
                    KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdARPG_SetData(_doc, ARPG_TZ_MainData, arpgRooms, aRPGFlats, _collection.ToArray()));
            }
            catch (Exception ex)
            {
                HtmlOutput.PrintError(ex);
            }
        }

        /// <summary>
        /// Реализация: Добавить новый тип квартир
        /// </summary>
        private void AddNewFlatData() => _collection.Add(new ARPG_TZ_FlatData());

        /// <summary>
        /// Реализация: Сортировать по тексту
        /// </summary>
        private void SortByText(string property)
        {
            using (ARPG_TZ_FlatDataColl.DeferRefresh())
            {
                ARPG_TZ_FlatDataColl.SortDescriptions.Clear();
                ARPG_TZ_FlatDataColl.SortDescriptions.Add(new SortDescription(property, ListSortDirection.Ascending));
            }
        }

        /// <summary>
        /// Реализация: Удалить тип квартир
        /// </summary>
        private void DeleteFlatData(ARPG_TZ_FlatData item)
        {
            UserDialog td = new UserDialog("KPLN: Подтверди действие", $"Подтверди удаление диапазона \"{item.TZRangeName}\"?");

            if ((bool)td.ShowDialog() && item != null && _collection.Contains(item))
                _collection.Remove(item);
        }
        #endregion

        /// <summary>
        /// Получить коллекцию ARPG_Room для проекта
        /// </summary>
        /// <returns></returns>
        private ARPG_Room[] ARPG_RoomsFromDoc(bool presetFlatCode)
        {
            FilteredElementCollector collector = new FilteredElementCollector(_doc);
            DesignOption[] designOptions = collector
                .OfClass(typeof(DesignOption))
                .Cast<DesignOption>()
                .ToArray();

            ARPG_Room[] arpgRooms;
            if (designOptions.Any())
            {
                AR_PyatnGraph_SelectDO selDO = new AR_PyatnGraph_SelectDO(designOptions);
                if ((bool)selDO.ShowDialog())
                    arpgRooms = ARPG_Room.Get_ARPG_Rooms(_doc, ARPG_TZ_MainData, selDO.SelARPGDesignOpt, presetFlatCode);
                else
                    return null;
            }
            else
                arpgRooms = ARPG_Room.Get_ARPG_Rooms(_doc, ARPG_TZ_MainData, new ARPG_DesOptEntity() { ARPG_DesignOptionId = -1 }, presetFlatCode);

            if (ARPG_Room.ErrorDict_Room.Keys.Count != 0)
            {
                HtmlOutput.PrintMsgDict("ОШИБКА", MessageType.Critical, ARPG_Room.ErrorDict_Room);

                MessageBox.Show(
                    $"Запуск невозможен. Список критических ошибок - выведен отдельным окном",
                    "Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);

                return null;
            }

            return arpgRooms;
        }

        private void OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Add && e.NewItems != null)
            {
                foreach (ARPG_TZ_FlatData item in e.NewItems)
                    SubscribeToItem(item);
            }
            else if (e.Action == NotifyCollectionChangedAction.Remove && e.OldItems != null)
            {
                foreach (ARPG_TZ_FlatData item in e.OldItems)
                    UnsubscribeFromItem(item);
            }
            else if (e.Action == NotifyCollectionChangedAction.Replace)
            {
                if (e.OldItems != null)
                    foreach (ARPG_TZ_FlatData it in e.OldItems) UnsubscribeFromItem(it);
                if (e.NewItems != null)
                    foreach (ARPG_TZ_FlatData it in e.NewItems) SubscribeToItem(it);
            }
            else if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                // коллекция была очищена/перезагружена: отпишемся от всего, затем подпишемся на текущие элементы
                foreach (var it in _subscribedItems.ToArray())
                    UnsubscribeFromItem(it);

                foreach (var it in _collection) // возможно новые элементы уже в _collection
                    SubscribeToItem(it);
            }

            UpdateSumPercent();
        }

        private void SubscribeToItem(ARPG_TZ_FlatData item)
        {
            if (item == null) return;

            if (_subscribedItems.Contains(item)) return;

            item.PropertyChanged += OnItemPropertyChanged;
            _subscribedItems.Add(item);
        }

        private void UnsubscribeFromItem(ARPG_TZ_FlatData item)
        {
            if (item == null) return;
            if (!_subscribedItems.Contains(item)) return;

            item.PropertyChanged -= OnItemPropertyChanged;
            _subscribedItems.Remove(item);
        }

        private void OnItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // реагируем только на FlatPercent (можно добавить другие свойства)
            if (e.PropertyName == nameof(ARPG_TZ_FlatData.TZPercent))
                UpdateSumPercent();
        }

        private void UpdateSumPercent()
        {
            double tempSum = 0;
            foreach (var item in _collection)
            {
                if (double.TryParse(item.TZPercent, out double val))
                    tempSum += val;
            }

            SumPercent = tempSum;
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
