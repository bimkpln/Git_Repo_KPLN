using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public class ApartmentPresetPanelContext
    {
        public List<ApartmentPlanPresetOption> Plans { get; set; }

        public string ActivePlanName { get; set; }

        public bool IsDataStale { get; set; }

        public Func<string, ApartmentPlanPresetOption> ResolvePlanData { get; set; }

        public ApartmentPresetPanelContext()
        {
            Plans = new List<ApartmentPlanPresetOption>();
        }

        public ApartmentPresetPanelContext CloneSnapshot()
        {
            ApartmentPresetPanelContext result = new ApartmentPresetPanelContext();
            result.ActivePlanName = ActivePlanName;
            result.IsDataStale = IsDataStale;

            if (Plans != null)
            {
                foreach (ApartmentPlanPresetOption plan in Plans)
                    result.Plans.Add(plan != null ? plan.Clone() : null);
            }

            return result;
        }
    }

    public class ApartmentPlanPresetOption
    {
        public string PlanName { get; set; }

        public string LowerConstraintText { get; set; }
        public string UpperConstraintText { get; set; }

        public List<int> WallThicknesses { get; set; }
        public Dictionary<int, List<string>> WallTypeOptionsByThickness { get; set; }

        public List<string> WindowTypeOptions { get; set; }
        public bool HasWindowMarkers { get; set; }

        public List<string> ShaftWallTypeOptions { get; set; }
        public bool HasShaftWallMarkers { get; set; }

        public List<string> RoomCategories { get; set; }

        public List<ApartmentDoorRequirementOption> DoorRequirements { get; set; }
        public Dictionary<string, List<string>> DoorTypeOptionsByRequirementKey { get; set; }

        public bool IsResolved { get; set; }

        public ApartmentPlanPresetOption()
        {
            WallThicknesses = new List<int>();
            WallTypeOptionsByThickness = new Dictionary<int, List<string>>();
            WindowTypeOptions = new List<string>();
            ShaftWallTypeOptions = new List<string>();

            RoomCategories = new List<string>();

            DoorRequirements = new List<ApartmentDoorRequirementOption>();
            DoorTypeOptionsByRequirementKey = new Dictionary<string, List<string>>();

            UpperConstraintText = "Неприсоединённая";
            IsResolved = false;
        }

        public ApartmentPlanPresetOption Clone()
        {
            ApartmentPlanPresetOption result = new ApartmentPlanPresetOption();
            result.PlanName = PlanName;
            result.LowerConstraintText = LowerConstraintText;
            result.UpperConstraintText = UpperConstraintText;
            result.WallThicknesses = WallThicknesses != null ? new List<int>(WallThicknesses) : new List<int>();
            result.WallTypeOptionsByThickness = CloneDictionaryList(WallTypeOptionsByThickness);
            result.WindowTypeOptions = WindowTypeOptions != null ? new List<string>(WindowTypeOptions) : new List<string>();
            result.HasWindowMarkers = HasWindowMarkers;
            result.ShaftWallTypeOptions = ShaftWallTypeOptions != null ? new List<string>(ShaftWallTypeOptions) : new List<string>();
            result.HasShaftWallMarkers = HasShaftWallMarkers;
            result.RoomCategories = RoomCategories != null ? new List<string>(RoomCategories) : new List<string>();
            result.DoorRequirements = DoorRequirements != null
                ? DoorRequirements.Select(x => x != null ? x.Clone() : null).ToList()
                : new List<ApartmentDoorRequirementOption>();
            result.DoorTypeOptionsByRequirementKey = CloneDictionaryList(DoorTypeOptionsByRequirementKey);
            result.IsResolved = IsResolved;
            return result;
        }

        private static Dictionary<TKey, List<TValue>> CloneDictionaryList<TKey, TValue>(Dictionary<TKey, List<TValue>> source)
        {
            Dictionary<TKey, List<TValue>> result = new Dictionary<TKey, List<TValue>>();
            if (source == null)
                return result;

            foreach (KeyValuePair<TKey, List<TValue>> kvp in source)
                result[kvp.Key] = kvp.Value != null ? new List<TValue>(kvp.Value) : new List<TValue>();

            return result;
        }
    }

    public class ApartmentDoorRequirementOption
    {
        public string RoomCategory { get; set; }

        public string DoorTypeName2D { get; set; }

        public int WidthMm { get; set; }

        public bool IsEntranceDoor { get; set; }

        public string Key
        {
            get { return BuildKey(RoomCategory, DoorTypeName2D, WidthMm, IsEntranceDoor); }
        }

        public string DisplayLabel
        {
            get
            {
                if (IsEntranceDoor)
                    return "Дверь входная (" + WidthMm + ")";

                string room = string.IsNullOrWhiteSpace(RoomCategory) ? "Без помещения" : RoomCategory;
                return "Дверь [" + room + "] (" + WidthMm + ")";
            }
        }

        public static string BuildKey(string roomCategory, string doorTypeName2D, int widthMm)
        {
            return BuildKey(roomCategory, doorTypeName2D, widthMm, false);
        }

        public static string BuildKey(string roomCategory, string doorTypeName2D, int widthMm, bool isEntranceDoor)
        {
            return (isEntranceDoor ? "ENTRY" : "ROOM") + "|" + (roomCategory ?? "") + "|" + (doorTypeName2D ?? "") + "|" + widthMm;
        }

        public ApartmentDoorRequirementOption Clone()
        {
            return new ApartmentDoorRequirementOption
            {
                RoomCategory = RoomCategory,
                DoorTypeName2D = DoorTypeName2D,
                WidthMm = WidthMm,
                IsEntranceDoor = IsEntranceDoor
            };
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
        public bool IsEntranceDoor { get; set; }

        public string Label
        {
            get
            {
                if (Kind == PresetSelectionKind.WallType)
                    return "Тип стены (" + (ThicknessMm.HasValue ? ThicknessMm.Value.ToString() : Key) + " мм)";

                string width = DoorWidthMm.HasValue ? DoorWidthMm.Value.ToString() : DoorTypeName2D ?? "";
                if (IsEntranceDoor)
                    return "Дверь входная (" + width + ")";

                string room = string.IsNullOrWhiteSpace(RoomCategory) ? "Без помещения" : RoomCategory;
                return "Дверь [" + room + "] (" + width + ")";
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

    internal class ApartmentPresetsVm : INotifyPropertyChanged
    {
        private const int ShaftWallTypeStorageKey = int.MinValue;

        public event Action<ApartmentPresetData> DataChanged;
        public event Action StateChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        private ApartmentPresetData _currentData;
        private ApartmentPresetPanelContext _context;
        private bool _isRefreshing;
        private bool _isDataStale;
        public ObservableCollection<ApartmentPlanPresetOption> Plans { get; private set; }
        public ObservableCollection<PresetSelectionVm> Assignments { get; private set; }
        public ObservableCollection<PresetSelectionVm> WallAssignments { get; private set; }
        public ObservableCollection<PresetSelectionVm> DoorAssignments { get; private set; }
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
                    NotifyDataChanged();
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
                    NotifyDataChanged();
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
                    NotifyDataChanged();
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
                    NotifyDataChanged();
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
                    NotifyDataChanged();
                }
            }
        }

        public ObservableCollection<string> WindowTypeOptions { get; private set; }

        private string _selectedWindowType;

        public string SelectedWindowType
        {
            get { return _selectedWindowType; }
            set
            {
                if (_selectedWindowType != value)
                {
                    _selectedWindowType = value;
                    OnPropertyChanged();
                    NotifyDataChanged();
                }
            }
        }

        private string _windowSillHeightText;

        public string WindowSillHeightText
        {
            get { return _windowSillHeightText; }
            set
            {
                if (_windowSillHeightText != value)
                {
                    _windowSillHeightText = value;
                    OnPropertyChanged();
                    NotifyDataChanged();
                }
            }
        }

        public bool HasWindowMarkers
        {
            get { return SelectedPlan != null && SelectedPlan.HasWindowMarkers; }
        }

        public ObservableCollection<string> ShaftWallTypeOptions { get; private set; }

        private string _selectedShaftWallType;

        public string SelectedShaftWallType
        {
            get { return _selectedShaftWallType; }
            set
            {
                if (_selectedShaftWallType != value)
                {
                    _selectedShaftWallType = value;
                    OnPropertyChanged();
                    NotifyDataChanged();
                }
            }
        }

        public bool HasShaftWallMarkers
        {
            get { return SelectedPlan != null && SelectedPlan.HasShaftWallMarkers; }
        }

        private string _statusText;

        public string StatusText
        {
            get { return _statusText; }
            private set
            {
                if (_statusText != value)
                {
                    _statusText = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool CanConvertTo3D
        {
            get
            {
                if (SelectedPlan == null || !SelectedPlan.IsResolved)
                    return false;

                if (_isDataStale)
                    return false;

                if (ParseInt(WallHeightText) <= 0)
                    return false;

                if (HasWindowMarkers)
                {
                    if (string.IsNullOrWhiteSpace(SelectedWindowType) ||
                        string.Equals(SelectedWindowType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        return false;

                    if (ParseInt(WindowSillHeightText, 900) <= 0)
                        return false;
                }

                if (HasShaftWallMarkers)
                {
                    if (string.IsNullOrWhiteSpace(SelectedShaftWallType) ||
                        string.Equals(SelectedShaftWallType, "Не выбрано", StringComparison.OrdinalIgnoreCase))
                        return false;
                }

                if (Assignments == null || Assignments.Count == 0)
                    return false;

                return Assignments.All(x =>
                    x != null &&
                    !string.IsNullOrWhiteSpace(x.SelectedValue) &&
                    !string.Equals(x.SelectedValue, "Не выбрано", StringComparison.OrdinalIgnoreCase));
            }
        }

        public ApartmentPresetsVm(ApartmentPresetData currentData, ApartmentPresetPanelContext context)
        {
            Plans = new ObservableCollection<ApartmentPlanPresetOption>();
            Assignments = new ObservableCollection<PresetSelectionVm>();
            WallAssignments = new ObservableCollection<PresetSelectionVm>();
            DoorAssignments = new ObservableCollection<PresetSelectionVm>();
            WindowTypeOptions = new ObservableCollection<string>();
            ShaftWallTypeOptions = new ObservableCollection<string>();

            _currentData = NormalizePresetData(currentData);
            ApplyContext(context, _currentData);
        }

        public void ApplyContext(ApartmentPresetPanelContext context, ApartmentPresetData currentData = null)
        {
            ApartmentPresetData dataToKeep = currentData != null
                ? currentData.Clone()
                : BuildData();

            _isRefreshing = true;

            try
            {
                _currentData = NormalizePresetData(dataToKeep);
                _context = context ?? new ApartmentPresetPanelContext();
                _isDataStale = _context.IsDataStale;

                Plans.Clear();
                foreach (ApartmentPlanPresetOption plan in _context.Plans)
                    Plans.Add(plan);

                BaseOffsetText = (_currentData.BaseOffset).ToString();
                WallHeightText = (_currentData.WallHeight > 0 ? _currentData.WallHeight : 3000).ToString();
                WindowSillHeightText = (_currentData.WindowSillHeight > 0 ? _currentData.WindowSillHeight : 900).ToString();

                ApartmentPlanPresetOption selected = ResolveInitialPlan();
                if (!ReferenceEquals(_selectedPlan, selected))
                {
                    _selectedPlan = selected;
                    EnsureSelectedPlanResolved();
                    OnPropertyChanged(nameof(SelectedPlan));
                    RefreshPlanDependentFields();
                }
                else
                {
                    EnsureSelectedPlanResolved();
                    RefreshPlanDependentFields();
                }

                if (SelectedPlan == null)
                {
                    LowerConstraintText = "";
                    UpperConstraintText = "Неприсоединённая";
                }
            }
            finally
            {
                _isRefreshing = false;
            }

            NotifyDataChanged();
        }

        public void MarkDataStale()
        {
            _isDataStale = true;
            if (_context != null)
                _context.IsDataStale = true;
            UpdateStateText();
            StateChanged?.Invoke();
            OnPropertyChanged(nameof(CanConvertTo3D));
        }

        public ApartmentPresetPanelContext GetContextSnapshot()
        {
            return _context != null
                ? _context.CloneSnapshot()
                : new ApartmentPresetPanelContext();
        }

        public ApartmentPresetData BuildData()
        {
            if (Assignments == null || Assignments.Count == 0)
            {
                ApartmentPresetData preserved = NormalizePresetData(_currentData);
                preserved.SelectedPlanName = SelectedPlan != null
                    ? SelectedPlan.PlanName
                    : preserved.SelectedPlanName;
                preserved.LowerConstraint = LowerConstraintText;
                preserved.UpperConstraint = UpperConstraintText;
                preserved.BaseOffset = ParseInt(BaseOffsetText);
                preserved.WallHeight = ParseInt(WallHeightText, 3000);
                preserved.WindowType = !string.IsNullOrWhiteSpace(SelectedWindowType) ? SelectedWindowType : preserved.WindowType;
                preserved.WindowSillHeight = ParseInt(WindowSillHeightText, 900);
                SetPresetShaftWallType(preserved, !string.IsNullOrWhiteSpace(SelectedShaftWallType) ? SelectedShaftWallType : GetPresetShaftWallType(preserved));
                return preserved;
            }

            Dictionary<int, string> wallTypeByThickness = Assignments != null
                ? Assignments
                    .Where(x => x.Kind == PresetSelectionKind.WallType && x.ThicknessMm.HasValue)
                    .GroupBy(x => x.ThicknessMm.Value)
                    .ToDictionary(
                        x => x.Key,
                        x => !string.IsNullOrWhiteSpace(x.First().SelectedValue)
                            ? x.First().SelectedValue
                            : "Не выбрано")
                : new Dictionary<int, string>();

            Dictionary<string, string> doorsByRoomCategory = Assignments != null
                ? Assignments
                    .Where(x => x.Kind == PresetSelectionKind.Door && !string.IsNullOrWhiteSpace(x.Key))
                    .GroupBy(x => x.Key)
                    .ToDictionary(
                        x => x.Key,
                        x => !string.IsNullOrWhiteSpace(x.First().SelectedValue)
                            ? x.First().SelectedValue
                            : "Не выбрано")
                : new Dictionary<string, string>();

            ApartmentPresetData result = new ApartmentPresetData
            {
                SelectedPlanName = SelectedPlan != null ? SelectedPlan.PlanName : "",
                LowerConstraint = LowerConstraintText,
                UpperConstraint = UpperConstraintText,
                BaseOffset = ParseInt(BaseOffsetText),
                WallHeight = ParseInt(WallHeightText, 3000),
                WallTypeByThickness = wallTypeByThickness,
                WindowType = !string.IsNullOrWhiteSpace(SelectedWindowType) ? SelectedWindowType : "Не выбрано",
                WindowSillHeight = ParseInt(WindowSillHeightText, 900),
                EntryDoor = "Не выбрано",
                BathroomDoor = "Не выбрано",
                RoomDoor = "Не выбрано",
                DoorsByRoomCategory = doorsByRoomCategory,
                FamilyPostProcessAction = _currentData != null
                    ? _currentData.FamilyPostProcessAction
                    : ApartmentFamilyPostProcessAction.Save2DFamiliesFromUnderlay
            };

            SetPresetShaftWallType(result, !string.IsNullOrWhiteSpace(SelectedShaftWallType) ? SelectedShaftWallType : "Не выбрано");

            return result;
        }

        public void SetFamilyPostProcessAction(ApartmentFamilyPostProcessAction action)
        {
            if (_currentData == null)
                _currentData = NormalizePresetData(null);

            _currentData.FamilyPostProcessAction = action;
            NotifyDataChanged();
        }

        private ApartmentPlanPresetOption ResolveInitialPlan()
        {
            if (Plans == null || Plans.Count == 0)
                return null;

            if (_currentData != null && !string.IsNullOrWhiteSpace(_currentData.SelectedPlanName))
            {
                ApartmentPlanPresetOption saved = Plans.FirstOrDefault(x =>
                    string.Equals(x.PlanName, _currentData.SelectedPlanName, StringComparison.OrdinalIgnoreCase));

                if (saved != null)
                    return saved;
            }

            if (_context != null && !string.IsNullOrWhiteSpace(_context.ActivePlanName))
            {
                ApartmentPlanPresetOption active = Plans.FirstOrDefault(x =>
                    string.Equals(x.PlanName, _context.ActivePlanName, StringComparison.OrdinalIgnoreCase));

                if (active != null)
                    return active;
            }

            return Plans.FirstOrDefault();
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
            _selectedPlan.WindowTypeOptions = resolved.WindowTypeOptions ?? new List<string>();
            _selectedPlan.HasWindowMarkers = resolved.HasWindowMarkers;
            _selectedPlan.ShaftWallTypeOptions = resolved.ShaftWallTypeOptions ?? new List<string>();
            _selectedPlan.HasShaftWallMarkers = resolved.HasShaftWallMarkers;
            _selectedPlan.RoomCategories = resolved.RoomCategories ?? new List<string>();
            _selectedPlan.DoorRequirements = resolved.DoorRequirements ?? new List<ApartmentDoorRequirementOption>();
            _selectedPlan.DoorTypeOptionsByRequirementKey = resolved.DoorTypeOptionsByRequirementKey ?? new Dictionary<string, List<string>>();
            _selectedPlan.IsResolved = true;
        }

        private void RefreshPlanDependentFields()
        {
            LowerConstraintText = SelectedPlan != null ? (SelectedPlan.LowerConstraintText ?? "") : "";

            UpperConstraintText = SelectedPlan != null ? (SelectedPlan.UpperConstraintText ?? "Неприсоединённая") : "Неприсоединённая";

            foreach (PresetSelectionVm assignment in Assignments)
                assignment.PropertyChanged -= Assignment_PropertyChanged;

            Assignments.Clear();
            WallAssignments.Clear();
            DoorAssignments.Clear();

            AddWallTypeAssignments();
            RefreshWindowFields();
            RefreshShaftWallFields();
            AddDoorAssignments();

            OnPropertyChanged(nameof(Assignments));
            OnPropertyChanged(nameof(WallAssignments));
            OnPropertyChanged(nameof(DoorAssignments));
            OnPropertyChanged(nameof(HasWindowMarkers));
            OnPropertyChanged(nameof(HasShaftWallMarkers));
            UpdateStateText();
        }

        private void RefreshWindowFields()
        {
            WindowTypeOptions.Clear();
            WindowTypeOptions.Add("Не выбрано");

            List<string> options = SelectedPlan != null && SelectedPlan.WindowTypeOptions != null
                ? SelectedPlan.WindowTypeOptions
                : new List<string>();

            foreach (string option in options
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x))
            {
                if (!WindowTypeOptions.Contains(option))
                    WindowTypeOptions.Add(option);
            }

            string savedWindowType = _currentData != null ? _currentData.WindowType : null;
            _selectedWindowType = !string.IsNullOrWhiteSpace(savedWindowType) &&
                                  WindowTypeOptions.Contains(savedWindowType)
                ? savedWindowType
                : "Не выбрано";

            OnPropertyChanged(nameof(SelectedWindowType));
            OnPropertyChanged(nameof(WindowTypeOptions));
        }

        private void RefreshShaftWallFields()
        {
            ShaftWallTypeOptions.Clear();
            ShaftWallTypeOptions.Add("Не выбрано");

            List<string> options = SelectedPlan != null && SelectedPlan.ShaftWallTypeOptions != null
                ? SelectedPlan.ShaftWallTypeOptions
                : new List<string>();

            foreach (string option in options
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .OrderBy(x => x))
            {
                if (!ShaftWallTypeOptions.Contains(option))
                    ShaftWallTypeOptions.Add(option);
            }

            string savedShaftWallType = GetPresetShaftWallType(_currentData);
            _selectedShaftWallType = !string.IsNullOrWhiteSpace(savedShaftWallType) &&
                                     ShaftWallTypeOptions.Contains(savedShaftWallType)
                ? savedShaftWallType
                : "Не выбрано";

            OnPropertyChanged(nameof(SelectedShaftWallType));
            OnPropertyChanged(nameof(ShaftWallTypeOptions));
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
                if (_currentData.WallTypeByThickness != null &&
                    _currentData.WallTypeByThickness.ContainsKey(thickness))
                {
                    savedWallType = _currentData.WallTypeByThickness[thickness];
                }

                vm.SelectedValue = !string.IsNullOrWhiteSpace(savedWallType) &&
                                   vm.Options.Contains(savedWallType)
                    ? savedWallType
                    : "Не выбрано";

                AddAssignment(vm);
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
                .GroupBy(x => ApartmentDoorRequirementOption.BuildKey(x.RoomCategory, x.DoorTypeName2D, x.WidthMm, x.IsEntranceDoor))
                .Select(x => x.First())
                .OrderBy(x => !x.IsEntranceDoor)
                .ThenBy(x => x.RoomCategory)
                .ThenBy(x => x.WidthMm)
                .ThenBy(x => x.DoorTypeName2D))
            {
                string key = ApartmentDoorRequirementOption.BuildKey(requirement.RoomCategory, requirement.DoorTypeName2D, requirement.WidthMm, requirement.IsEntranceDoor);

                PresetSelectionVm vm = new PresetSelectionVm
                {
                    Kind = PresetSelectionKind.Door,
                    Key = key,
                    RoomCategory = requirement.RoomCategory,
                    DoorTypeName2D = requirement.DoorTypeName2D,
                    DoorWidthMm = requirement.WidthMm,
                    IsEntranceDoor = requirement.IsEntranceDoor
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
                if (_currentData.DoorsByRoomCategory != null &&
                    _currentData.DoorsByRoomCategory.ContainsKey(key))
                {
                    savedDoor = _currentData.DoorsByRoomCategory[key];
                }

                vm.SelectedValue = !string.IsNullOrWhiteSpace(savedDoor) &&
                                   vm.Options.Contains(savedDoor)
                    ? savedDoor
                    : "Не выбрано";

                AddAssignment(vm);
            }
        }

        private void AddAssignment(PresetSelectionVm vm)
        {
            if (vm == null)
                return;

            vm.PropertyChanged += Assignment_PropertyChanged;
            Assignments.Add(vm);

            if (vm.Kind == PresetSelectionKind.WallType)
                WallAssignments.Add(vm);
            else if (vm.Kind == PresetSelectionKind.Door)
                DoorAssignments.Add(vm);
        }

        private void Assignment_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e == null || string.IsNullOrWhiteSpace(e.PropertyName) || e.PropertyName == nameof(PresetSelectionVm.SelectedValue))
                NotifyDataChanged();
        }

        private static int ParseInt(string text, int defaultValue = 0)
        {
            int value;
            return int.TryParse(text, out value) ? value : defaultValue;
        }

        private void NotifyDataChanged()
        {
            if (_isRefreshing)
                return;

            _currentData = NormalizePresetData(BuildData());
            UpdateStateText();

            DataChanged?.Invoke(_currentData.Clone());
            StateChanged?.Invoke();

            OnPropertyChanged(nameof(CanConvertTo3D));
        }

        private void UpdateStateText()
        {
            if (SelectedPlan == null)
            {
                StatusText = "Нажмите «Обновить данные» на открытом плане этажа.";
                return;
            }

            if (_isDataStale)
            {
                StatusText = "Данные устарели после размещения 2D-семейства. Нажмите «Обновить данные».";
                return;
            }

            if (!SelectedPlan.IsResolved)
            {
                StatusText = "Данные плана ещё не обновлены.";
                return;
            }

            if (Assignments == null || Assignments.Count == 0)
            {
                StatusText = "На плане не найдены размещённые 2D-семейства квартир.";
                return;
            }

            if (!CanConvertTo3D)
            {
                StatusText = "Заполните все найденные стены, шахты, двери и окна.";
                return;
            }

            StatusText = "Данные заполнены.";
        }

        private static ApartmentPresetData NormalizePresetData(ApartmentPresetData data)
        {
            ApartmentPresetData result = data != null
                ? data.Clone()
                : new ApartmentPresetData();

            if (result.WallTypeByThickness == null)
                result.WallTypeByThickness = new Dictionary<int, string>();

            if (result.DoorsByRoomCategory == null)
                result.DoorsByRoomCategory = new Dictionary<string, string>();

            if (string.IsNullOrWhiteSpace(result.UpperConstraint))
                result.UpperConstraint = "Неприсоединённая";

            if (result.WallHeight <= 0)
                result.WallHeight = 3000;

            if (string.IsNullOrWhiteSpace(result.WindowType))
                result.WindowType = "Не выбрано";

            if (result.WindowSillHeight <= 0)
                result.WindowSillHeight = 900;

            if (string.IsNullOrWhiteSpace(GetPresetShaftWallType(result)))
                SetPresetShaftWallType(result, "Не выбрано");

            if (string.IsNullOrWhiteSpace(result.EntryDoor))
                result.EntryDoor = "Не выбрано";

            if (string.IsNullOrWhiteSpace(result.BathroomDoor))
                result.BathroomDoor = "Не выбрано";

            if (string.IsNullOrWhiteSpace(result.RoomDoor))
                result.RoomDoor = "Не выбрано";

            return result;
        }

        private static string GetPresetShaftWallType(ApartmentPresetData data)
        {
            if (data == null)
                return "Не выбрано";

            string value;
            if (data.WallTypeByThickness != null &&
                data.WallTypeByThickness.TryGetValue(ShaftWallTypeStorageKey, out value) &&
                !string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            try
            {
                var prop = data.GetType().GetProperty("ShaftWallType");
                if (prop != null && prop.PropertyType == typeof(string))
                {
                    value = prop.GetValue(data, null) as string;
                    if (!string.IsNullOrWhiteSpace(value))
                        return value;
                }
            }
            catch
            {
            }

            return "Не выбрано";
        }

        private static void SetPresetShaftWallType(ApartmentPresetData data, string value)
        {
            if (data == null)
                return;

            if (string.IsNullOrWhiteSpace(value))
                value = "Не выбрано";

            if (data.WallTypeByThickness == null)
                data.WallTypeByThickness = new Dictionary<int, string>();

            data.WallTypeByThickness[ShaftWallTypeStorageKey] = value;

            try
            {
                var prop = data.GetType().GetProperty("ShaftWallType");
                if (prop != null && prop.PropertyType == typeof(string) && prop.CanWrite)
                    prop.SetValue(data, value, null);
            }
            catch
            {
            }
        }

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged;
            if (h != null)
                h(this, new PropertyChangedEventArgs(p));
        }
    }
}