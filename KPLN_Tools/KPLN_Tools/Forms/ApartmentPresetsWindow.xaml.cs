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
    public class ApartmentPresetWindowContext
    {
        public List<ApartmentPlanPresetOption> Plans { get; set; }

        public ApartmentPresetWindowContext()
        {
            Plans = new List<ApartmentPlanPresetOption>();
        }
    }

    public class ApartmentPlanPresetOption
    {
        public string PlanName { get; set; }
        public string LowerConstraintText { get; set; }
        public string UpperConstraintText { get; set; }

        public List<int> WallThicknesses { get; set; }
        public Dictionary<int, List<string>> WallTypeOptionsByThickness { get; set; }

        public List<string> RoomCategories { get; set; }

        public ApartmentPlanPresetOption()
        {
            WallThicknesses = new List<int>();
            WallTypeOptionsByThickness = new Dictionary<int, List<string>>();
            RoomCategories = new List<string>();
            UpperConstraintText = "Неприсоединённая";
        }
    }

    public enum PresetSelectionKind
    {
        WallType,
        Door
    }

    public class PresetSelectionVm : INotifyPropertyChanged
    {
        public PresetSelectionKind Kind { get; set; }

        /// Ключ элемента: для стены: толщина в мм, для двери: категория помещения

        public string Key { get; set; }

        public int? ThicknessMm { get; set; }
        public string RoomCategory { get; set; }

        public string Label
        {
            get
            {
                if (Kind == PresetSelectionKind.WallType)
                    return "Тип стены (" + (ThicknessMm.HasValue ? ThicknessMm.Value.ToString() : Key) + ")";

                return "Дверь (" + (RoomCategory ?? Key ?? "") + ")";
            }
        }

        public ObservableCollection<string> Options { get; private set; }

        private string _selectedValue;
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

    public partial class ApartmentPresetsWindow : Window
    {
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

        private void Vm_RequestClose()
        {
            DialogResult = false;
            Close();
        }

        private void Vm_RequestSave(ApartmentPresetData data)
        {
            ResultPresetData = data;
            DialogResult = true;
            Close();
        }

        private void DigitsOnly_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            e.Handled = e.Text.Any(x => !char.IsDigit(x));
        }

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
    }

    internal class ApartmentPresetsVm : INotifyPropertyChanged
    {
        public event Action RequestClose;
        public event Action<ApartmentPresetData> RequestSave;

        private readonly ApartmentPresetData _initialData;
        private readonly ApartmentPresetWindowContext _context;

        public ObservableCollection<ApartmentPlanPresetOption> Plans { get; private set; }

        public ObservableCollection<PresetSelectionVm> Assignments { get; private set; }

        private ApartmentPlanPresetOption _selectedPlan;
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

        public ICommand SaveCommand { get; private set; }
        public ICommand CloseCommand { get; private set; }

        public ApartmentPresetsVm(
            ApartmentPresetData currentData,
            ApartmentPresetWindowContext context)
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

        private void RefreshPlanDependentFields()
        {
            LowerConstraintText = SelectedPlan != null
                ? (SelectedPlan.LowerConstraintText ?? "")
                : "";

            UpperConstraintText = SelectedPlan != null
                ? (SelectedPlan.UpperConstraintText ?? "Неприсоединённая")
                : "Неприсоединённая";

            Assignments.Clear();

            AddWallTypeAssignments();
            AddDoorAssignments();

            OnPropertyChanged(nameof(Assignments));
        }

        private void AddWallTypeAssignments()
        {
            List<int> thicknesses = SelectedPlan != null && SelectedPlan.WallThicknesses != null
                ? SelectedPlan.WallThicknesses
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

        private void AddDoorAssignments()
        {
            List<string> roomCategories = SelectedPlan != null && SelectedPlan.RoomCategories != null
                ? SelectedPlan.RoomCategories
                : new List<string>();

            if (roomCategories.Count == 0)
                roomCategories.Add("Помещение");

            foreach (string category in roomCategories
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x))
            {
                PresetSelectionVm vm = new PresetSelectionVm
                {
                    Kind = PresetSelectionKind.Door,
                    Key = category,
                    RoomCategory = category
                };

                vm.Options.Add("Заглушка");

                string savedDoor = null;
                if (_initialData.DoorsByRoomCategory != null &&
                    _initialData.DoorsByRoomCategory.ContainsKey(category))
                {
                    savedDoor = _initialData.DoorsByRoomCategory[category];
                }

                vm.SelectedValue = !string.IsNullOrWhiteSpace(savedDoor) &&
                                   vm.Options.Contains(savedDoor)
                    ? savedDoor
                    : "Заглушка";

                Assignments.Add(vm);
            }
        }

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
                .Where(x => x.Kind == PresetSelectionKind.Door && !string.IsNullOrWhiteSpace(x.RoomCategory))
                .GroupBy(x => x.RoomCategory)
                .ToDictionary(
                    x => x.Key,
                    x => !string.IsNullOrWhiteSpace(x.First().SelectedValue)
                        ? x.First().SelectedValue
                        : "Заглушка");

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

        private static int ParseInt(string text, int defaultValue = 0)
        {
            int value;
            return int.TryParse(text, out value) ? value : defaultValue;
        }

        private void OnClose()
        {
            RequestClose?.Invoke();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }
}