using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    /// Контекст окна преднастроек.  Содержит список доступных планов, которые будут показаны в ComboBox.
    public class ApartmentPresetWindowContext
    {
        public List<ApartmentPlanPresetOption> Plans { get; set; }

        public ApartmentPresetWindowContext()
        {
            Plans = new List<ApartmentPlanPresetOption>();
        }
    }

    /// Описание одного доступного плана квартиры.
    /// Включает в себя:
    /// - имя плана;
    /// - текст зависимостей;
    /// - набор возможных толщин стен;
    /// - варианты типов стен по толщине;
    /// - список требований по дверям;
    /// - варианты типов дверей по ключу требования.
    public class ApartmentPlanPresetOption
    {
        public string PlanName { get; set; }

        public string LowerConstraintText { get; set; }
        public string UpperConstraintText { get; set; }

        public List<int> WallThicknesses { get; set; }
        public Dictionary<int, List<string>> WallTypeOptionsByThickness { get; set; }

        public List<string> RoomCategories { get; set; }

        public List<ApartmentDoorRequirementOption> DoorRequirements { get; set; }
        public Dictionary<string, List<string>> DoorTypeOptionsByRequirementKey { get; set; }

        public ApartmentPlanPresetOption()
        {
            WallThicknesses = new List<int>();
            WallTypeOptionsByThickness = new Dictionary<int, List<string>>();

            RoomCategories = new List<string>();

            DoorRequirements = new List<ApartmentDoorRequirementOption>();
            DoorTypeOptionsByRequirementKey = new Dictionary<string, List<string>>();

            UpperConstraintText = "Неприсоединённая";
        }
    }

    /// Описание требования к двери для конкретного плана. Ключ строится по имени 2D-типа двери и её ширине.
    public class ApartmentDoorRequirementOption
    {
        public string DoorTypeName2D { get; set; }

        public int WidthMm { get; set; }

        public string Key
        {
            get { return BuildKey(DoorTypeName2D, WidthMm); }
        }

        public static string BuildKey(string doorTypeName2D, int widthMm)
        {
            return (doorTypeName2D ?? "") + "|" + widthMm;
        }
    }

    /// Тип назначения в списке преднастроек:
    public enum PresetSelectionKind
    {
        WallType,
        Door
    }

    /// <summary>
    /// VM одной строки в блоке "Выбор типов". Используется как для выбора типа стены, так и для выбора типа двери.
    /// </summary>
    public class PresetSelectionVm : INotifyPropertyChanged
    {
        /// Вид назначения: стена или дверь.
        public PresetSelectionKind Kind { get; set; }

        /// Ключ назначения.
        /// Для стены обычно толщина в текстовом виде,
        /// для двери — составной ключ DoorTypeName2D|WidthMm.
        public string Key { get; set; }

        /// Толщина стены в мм, если строка относится к типу стены.
        public int? ThicknessMm { get; set; }

        /// Имя 2D-типа двери, если строка относится к двери.
        public string DoorTypeName2D { get; set; }

        /// Ширина двери в мм, если строка относится к двери.
        public int? DoorWidthMm { get; set; }

        /// Подпись строки для показа в интерфейсе.
        public string Label
        {
            get
            {
                if (Kind == PresetSelectionKind.WallType)
                    return "Тип стены (" + (ThicknessMm.HasValue ? ThicknessMm.Value.ToString() : Key) + " мм)";

                return "Дверь (" + (DoorTypeName2D ?? "") + ")";
            }
        }

        /// Доступные варианты выбора для ComboBox.
        public ObservableCollection<string> Options { get; private set; }
        private string _selectedValue;

        /// Выбранное пользователем значение.
        public string SelectedValue
        {
            get { return _selectedValue; }
            set
            {
                if (_selectedValue != value)
                {
                    _selectedValue = value;
                    OnPropertyChanged();
                }
            }
        }

        public PresetSelectionVm()
        {
            Options = new ObservableCollection<string>();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }

    /// Окно преднастроек квартиры.
    public partial class ApartmentPresetsWindow : Window
    {
        /// Результат работы окна после успешного сохранения.
        public ApartmentPresetData ResultPresetData { get; private set; }

        public ApartmentPresetsWindow(
            ApartmentPresetData currentData,
            ApartmentPresetWindowContext context)
        {
            InitializeComponent();

            var vm = new ApartmentPresetsVm(currentData, context);
            vm.RequestClose += Vm_RequestClose;
            vm.RequestSave += Vm_RequestSave;

            DataContext = vm;
        }

        /// Разрешаем ввод только цифр в числовые поля.
        private void DigitsOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(x => !char.IsDigit(x));
        }

        /// Запрещаем вставку текста, если в нём есть нецифровые символы.
        private void DigitsOnly_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (!e.DataObject.GetDataPresent(typeof(string)))
            {
                e.CancelCommand();
                return;
            }

            string text = e.DataObject.GetData(typeof(string)) as string;
            if (string.IsNullOrWhiteSpace(text) || text.Any(x => !char.IsDigit(x)))
                e.CancelCommand();
        }

        /// Закрытие окна без сохранения.
        private void Vm_RequestClose()
        {
            DialogResult = false;
            Close();
        }

        /// Закрытие окна с сохранением результата.
        private void Vm_RequestSave(ApartmentPresetData data)
        {
            ResultPresetData = data;
            DialogResult = true;
            Close();
        }
    }

    /// ViewModel окна преднастроек квартиры.
    /// - список планов;
    /// - выбранный план;
    /// - отображение зависимостей;
    /// - формирование динамического списка назначений;
    /// - сбор результата при сохранении.
    internal class ApartmentPresetsVm : INotifyPropertyChanged
    {
        public event Action RequestClose;
        public event Action<ApartmentPresetData> RequestSave;
        public event PropertyChangedEventHandler PropertyChanged;

        /// Исходные данные пресета. Используются для начальной инициализации значений.
        private readonly ApartmentPresetData _initialData;
        /// Контекст окна со списком доступных планов и опций.
        private readonly ApartmentPresetWindowContext _context;
        /// Список доступных планов для выбора.
        public ObservableCollection<ApartmentPlanPresetOption> Plans { get; private set; }
        /// Динамический список строк выбора: сначала стены, потом двери.
        public ObservableCollection<PresetSelectionVm> Assignments { get; private set; }
        private ApartmentPlanPresetOption _selectedPlan;

        /// Выбранный план. При изменении пересобираются все зависимые поля и назначения.
        public ApartmentPlanPresetOption SelectedPlan
        {
            get { return _selectedPlan; }
            set
            {
                if (_selectedPlan != value)
                {
                    _selectedPlan = value;
                    OnPropertyChanged();
                    RefreshPlanDependentFields();
                }
            }
        }

        private string _lowerConstraintText;

        /// Текст нижней зависимости для выбранного плана.
        public string LowerConstraintText
        {
            get { return _lowerConstraintText; }
            set
            {
                if (_lowerConstraintText != value)
                {
                    _lowerConstraintText = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _upperConstraintText;

        /// Текст верхней зависимости для выбранного плана.
        public string UpperConstraintText
        {
            get { return _upperConstraintText; }
            set
            {
                if (_upperConstraintText != value)
                {
                    _upperConstraintText = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _baseOffsetText;

        /// Смещение снизу в текстовом виде. Хранится строкой, потому что связано напрямую с TextBox.
        public string BaseOffsetText
        {
            get { return _baseOffsetText; }
            set
            {
                if (_baseOffsetText != value)
                {
                    _baseOffsetText = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _wallHeightText;

        /// Неприсоединённая высота в текстовом виде. Хранится строкой, потому что связано напрямую с TextBox.
        public string WallHeightText
        {
            get { return _wallHeightText; }
            set
            {
                if (_wallHeightText != value)
                {
                    _wallHeightText = value;
                    OnPropertyChanged();
                }
            }
        }

        /// Команда сохранения результата.
        public ICommand SaveCommand { get; private set; }

        /// Команда закрытия окна без сохранения.
        public ICommand CloseCommand { get; private set; }

        public ApartmentPresetsVm(ApartmentPresetData currentData, ApartmentPresetWindowContext context)
        {
            _initialData = currentData != null ? currentData.Clone() : new ApartmentPresetData();
            _context = context ?? new ApartmentPresetWindowContext();

            Plans = new ObservableCollection<ApartmentPlanPresetOption>();
            Assignments = new ObservableCollection<PresetSelectionVm>();

            foreach (ApartmentPlanPresetOption plan in _context.Plans)
                Plans.Add(plan);

            BaseOffsetText = (_initialData.BaseOffset).ToString();
            WallHeightText = (_initialData.WallHeight > 0 ? _initialData.WallHeight : 3000).ToString();

            SaveCommand = new RelayCommand(OnSave);
            CloseCommand = new RelayCommand(OnClose);

            SelectedPlan = Plans.FirstOrDefault(x => x.PlanName == _initialData.SelectedPlanName)
                           ?? Plans.FirstOrDefault();

            if (SelectedPlan == null)
            {
                LowerConstraintText = "";
                UpperConstraintText = "Неприсоединённая";
            }
        }

        /// Полностью обновляет все данные, зависящие от выбранного плана:
        /// - тексты зависимостей;
        /// - список назначений по стенам;
        /// - список назначений по дверям.
        private void RefreshPlanDependentFields()
        {
            LowerConstraintText = SelectedPlan != null ? (SelectedPlan.LowerConstraintText ?? "") : "";

            UpperConstraintText = SelectedPlan != null ? (SelectedPlan.UpperConstraintText ?? "Неприсоединённая") : "Неприсоединённая";

            Assignments.Clear();

            AddWallTypeAssignments();
            AddDoorAssignments();

            OnPropertyChanged(nameof(Assignments));
        }

        /// Формирует строки выбора для типов стен. Каждая уникальная толщина стены превращается в отдельную строку.
        private void AddWallTypeAssignments()
        {
            List<int> thicknesses = SelectedPlan != null && SelectedPlan.WallThicknesses != null ? SelectedPlan.WallThicknesses
                : new List<int>();

            foreach (int thickness in thicknesses.Distinct().OrderBy(x => x))
            {
                PresetSelectionVm vm = new PresetSelectionVm
                {
                    Kind = PresetSelectionKind.WallType,
                    Key = thickness.ToString(),
                    ThicknessMm = thickness
                };

                vm.Options.Add("Не выбрано");

                List<string> options;
                if (SelectedPlan != null &&
                    SelectedPlan.WallTypeOptionsByThickness != null &&
                    SelectedPlan.WallTypeOptionsByThickness.TryGetValue(thickness, out options) &&
                    options != null)
                {
                    foreach (string option in options
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .OrderBy(x => x))
                    {
                        if (!vm.Options.Contains(option))
                            vm.Options.Add(option);
                    }
                }

                string savedWallType = null;
                if (_initialData.WallTypeByThickness != null &&
                    _initialData.WallTypeByThickness.ContainsKey(thickness))
                {
                    savedWallType = _initialData.WallTypeByThickness[thickness];
                }

                vm.SelectedValue = !string.IsNullOrWhiteSpace(savedWallType) &&
                                   vm.Options.Contains(savedWallType)
                    ? savedWallType
                    : "Не выбрано";

                Assignments.Add(vm);
            }
        }

        /// Формирует строки выбора для дверей. Каждое уникальное требование двери превращается в отдельную строку.
        private void AddDoorAssignments()
        {
            List<ApartmentDoorRequirementOption> requirements =
                SelectedPlan != null && SelectedPlan.DoorRequirements != null
                    ? SelectedPlan.DoorRequirements
                    : new List<ApartmentDoorRequirementOption>();

            foreach (ApartmentDoorRequirementOption requirement in requirements
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.DoorTypeName2D) && x.WidthMm > 0)
                .GroupBy(x => ApartmentDoorRequirementOption.BuildKey(x.DoorTypeName2D, x.WidthMm))
                .Select(x => x.First())
                .OrderBy(x => x.DoorTypeName2D))
            {
                string key = ApartmentDoorRequirementOption.BuildKey(requirement.DoorTypeName2D, requirement.WidthMm);

                PresetSelectionVm vm = new PresetSelectionVm
                {
                    Kind = PresetSelectionKind.Door,
                    Key = key,
                    DoorTypeName2D = requirement.DoorTypeName2D,
                    DoorWidthMm = requirement.WidthMm
                };

                vm.Options.Add("Не выбрано");

                List<string> options;
                if (SelectedPlan != null &&
                    SelectedPlan.DoorTypeOptionsByRequirementKey != null &&
                    SelectedPlan.DoorTypeOptionsByRequirementKey.TryGetValue(key, out options) &&
                    options != null)
                {
                    foreach (string option in options
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct()
                        .OrderBy(x => x))
                    {
                        if (!vm.Options.Contains(option))
                            vm.Options.Add(option);
                    }
                }

                string savedDoor = null;
                if (_initialData.DoorsByRoomCategory != null &&
                    _initialData.DoorsByRoomCategory.ContainsKey(key))
                {
                    savedDoor = _initialData.DoorsByRoomCategory[key];
                }

                vm.SelectedValue = !string.IsNullOrWhiteSpace(savedDoor) &&
                                   vm.Options.Contains(savedDoor)
                    ? savedDoor
                    : "Не выбрано";

                Assignments.Add(vm);
            }
        }

        /// Собирает результат из текущего состояния интерфейса и передаёт его наружу через событие RequestSave.
        private void OnSave()
        {
            Dictionary<int, string> wallTypeByThickness = Assignments
                .Where(x => x.Kind == PresetSelectionKind.WallType && x.ThicknessMm.HasValue)
                .GroupBy(x => x.ThicknessMm.Value)
                .ToDictionary(
                    x => x.Key,
                    x => !string.IsNullOrWhiteSpace(x.First().SelectedValue)
                        ? x.First().SelectedValue
                        : "Не выбрано");

            Dictionary<string, string> doorsByRoomCategory = Assignments
                .Where(x => x.Kind == PresetSelectionKind.Door && !string.IsNullOrWhiteSpace(x.Key))
                .GroupBy(x => x.Key)
                .ToDictionary(
                    x => x.Key,
                    x => !string.IsNullOrWhiteSpace(x.First().SelectedValue)
                        ? x.First().SelectedValue
                        : "Не выбрано");

            ApartmentPresetData data = new ApartmentPresetData
            {
                SelectedPlanName = SelectedPlan != null ? SelectedPlan.PlanName : "",
                LowerConstraint = LowerConstraintText,
                UpperConstraint = UpperConstraintText,
                BaseOffset = ParseInt(BaseOffsetText),
                WallHeight = ParseInt(WallHeightText, 3000),
                WallTypeByThickness = wallTypeByThickness,
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = doorsByRoomCategory
            };

            RequestSave?.Invoke(data);
        }

        /// Закрытие окна без сохранения.
        private void OnClose()
        {
            RequestClose?.Invoke();
        }

        /// Парсинг int из текстового значения с дефолтом.
        private static int ParseInt(string text, int defaultValue = 0)
        {
            int value;
            return int.TryParse(text, out value) ? value : defaultValue;
        }

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }
}