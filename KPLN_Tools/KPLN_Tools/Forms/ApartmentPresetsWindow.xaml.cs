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

        public Func<string, ApartmentPlanPresetOption> ResolvePlanData { get; set; }

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

        public List<ApartmentDoorRequirementOption> DoorRequirements { get; set; }
        public Dictionary<string, List<string>> DoorTypeOptionsByRequirementKey { get; set; }

        public bool IsResolved { get; set; }

        public ApartmentPlanPresetOption()
        {
            WallThicknesses = new List<int>();
            WallTypeOptionsByThickness = new Dictionary<int, List<string>>();

            RoomCategories = new List<string>();

            DoorRequirements = new List<ApartmentDoorRequirementOption>();
            DoorTypeOptionsByRequirementKey = new Dictionary<string, List<string>>();

            UpperConstraintText = "Неприсоединённая";
            IsResolved = false;
        }
    }

    public class ApartmentDoorRequirementOption
    {
        public string RoomCategory { get; set; }

        public string DoorTypeName2D { get; set; }

        public int WidthMm { get; set; }

        public string Key
        {
            get { return BuildKey(RoomCategory, DoorTypeName2D, WidthMm); }
        }

        public string DisplayLabel
        {
            get
            {
                string room = string.IsNullOrWhiteSpace(RoomCategory) ? "Без помещения" : RoomCategory;
                return "Дверь [" + room + "] (" + WidthMm + ")";
            }
        }

        public static string BuildKey(string roomCategory, string doorTypeName2D, int widthMm)
        {
            return (roomCategory ?? "") + "|" + (doorTypeName2D ?? "") + "|" + widthMm;
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

        public string Key { get; set; }
        public int? ThicknessMm { get; set; }
        public string DoorTypeName2D { get; set; }
        public string RoomCategory { get; set; }
        public int? DoorWidthMm { get; set; }

        public string Label
        {
            get
            {
                if (Kind == PresetSelectionKind.WallType)
                    return "Тип стены (" + (ThicknessMm.HasValue ? ThicknessMm.Value.ToString() : Key) + " мм)";

                string room = string.IsNullOrWhiteSpace(RoomCategory) ? "Без помещения" : RoomCategory;
                return "Дверь [" + room + "] (" + (DoorWidthMm.HasValue ? DoorWidthMm.Value.ToString() : DoorTypeName2D ?? "") + ")";
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
    }

    internal class ApartmentPresetsVm : INotifyPropertyChanged
    {
        public event Action RequestClose;
        public event Action<ApartmentPresetData> RequestSave;
        public event PropertyChangedEventHandler PropertyChanged;

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
                    EnsureSelectedPlanResolved();
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

        private void EnsureSelectedPlanResolved()
        {
            if (_selectedPlan == null)
                return;

            if (_selectedPlan.IsResolved)
                return;

            if (_context == null || _context.ResolvePlanData == null)
                return;

            ApartmentPlanPresetOption resolved = _context.ResolvePlanData(_selectedPlan.PlanName);
            if (resolved == null)
                return;

            _selectedPlan.LowerConstraintText = resolved.LowerConstraintText;
            _selectedPlan.UpperConstraintText = resolved.UpperConstraintText;
            _selectedPlan.WallThicknesses = resolved.WallThicknesses ?? new List<int>();
            _selectedPlan.WallTypeOptionsByThickness = resolved.WallTypeOptionsByThickness ?? new Dictionary<int, List<string>>();
            _selectedPlan.RoomCategories = resolved.RoomCategories ?? new List<string>();
            _selectedPlan.DoorRequirements = resolved.DoorRequirements ?? new List<ApartmentDoorRequirementOption>();
            _selectedPlan.DoorTypeOptionsByRequirementKey = resolved.DoorTypeOptionsByRequirementKey ?? new Dictionary<string, List<string>>();
            _selectedPlan.IsResolved = true;
        }

        private void RefreshPlanDependentFields()
        {
            LowerConstraintText = SelectedPlan != null ? (SelectedPlan.LowerConstraintText ?? "") : "";

            UpperConstraintText = SelectedPlan != null ? (SelectedPlan.UpperConstraintText ?? "Неприсоединённая") : "Неприсоединённая";

            Assignments.Clear();

            AddWallTypeAssignments();
            AddDoorAssignments();

            OnPropertyChanged(nameof(Assignments));
        }

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

        private void AddDoorAssignments()
        {
            List<ApartmentDoorRequirementOption> requirements =
                SelectedPlan != null && SelectedPlan.DoorRequirements != null
                    ? SelectedPlan.DoorRequirements
                    : new List<ApartmentDoorRequirementOption>();

            foreach (ApartmentDoorRequirementOption requirement in requirements
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.DoorTypeName2D) && x.WidthMm > 0)
                .GroupBy(x => ApartmentDoorRequirementOption.BuildKey(x.RoomCategory, x.DoorTypeName2D, x.WidthMm))
                .Select(x => x.First())
                .OrderBy(x => x.RoomCategory)
                .ThenBy(x => x.WidthMm)
                .ThenBy(x => x.DoorTypeName2D))
            {
                string key = ApartmentDoorRequirementOption.BuildKey(requirement.RoomCategory, requirement.DoorTypeName2D, requirement.WidthMm);

                PresetSelectionVm vm = new PresetSelectionVm
                {
                    Kind = PresetSelectionKind.Door,
                    Key = key,
                    RoomCategory = requirement.RoomCategory,
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
                DoorsByRoomCategory = doorsByRoomCategory,
                FamilyPostProcessAction = _initialData != null ? _initialData.FamilyPostProcessAction : ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            RequestSave?.Invoke(data);
        }

        private void OnClose()
        {
            RequestClose?.Invoke();
        }

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