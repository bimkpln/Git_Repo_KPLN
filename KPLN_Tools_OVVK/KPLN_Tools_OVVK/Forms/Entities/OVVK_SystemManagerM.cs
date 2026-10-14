using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
using KPLN_Tools_OVVK.Common.OVVK_System;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Tools_OVVK.Forms.Entities
{
    public sealed class OVVK_SystemManagerM : INotifyPropertyChanged, IJsonSerializable
    {
        private static readonly BuiltInCategory[] _familyCategories =
        {
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_PipeAccessory,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_Sprinklers,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_GenericModel,
        };

        private static readonly BuiltInCategory[] _curveCategories =
        {
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_FlexDuctCurves,
            BuiltInCategory.OST_DuctInsulations,
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_FlexPipeCurves,
            BuiltInCategory.OST_PipeInsulations,
        };

        private static readonly string[] _excludedFamilyPrefixes =
        {
            "500_", "501_", "502_", "503_", "ASML_О_Отверстие_", "ClashPoint",
        };

        private string _parameterName = string.Empty;
        private string _sysNameSeparator = "/";
        private string _operationMessage = string.Empty;
        private readonly Dictionary<string, List<ElementId>> _systemElementIds = new Dictionary<string, List<ElementId>>();
        private readonly Dictionary<OVVK_SystemManagerRowKind, List<ElementId>> _errorElementIds = new Dictionary<OVVK_SystemManagerRowKind, List<ElementId>>();
        private bool _isBusy;
        private bool _isDocumentActive = true;


        [JsonConstructor]
        public OVVK_SystemManagerM()
        {
        }

        public OVVK_SystemManagerM(Document doc)
        {
            CurrentDoc = doc;
            DocumentTitle = doc.Title;

            // Настройки другого проекта не должны подменять выбор параметра.
            ModelPath centralPath = doc.IsWorkshared ? doc.GetWorksharingCentralModelPath() : null;
            string documentPath = centralPath == null
                ? doc.PathName
                : ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath);
            DocumentKey = string.IsNullOrEmpty(documentPath)
                ? Guid.NewGuid().ToString()
                : doc.Application.VersionNumber + "|" + documentPath;

            Refresh();
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [JsonIgnore]
        public Document CurrentDoc { get; private set; }

        public string DocumentKey { get; set; }

        [JsonIgnore]
        public string DocumentTitle { get; private set; }

        [JsonIgnore]
        public List<Element> ElementColl { get; private set; } = new List<Element>();

        [JsonIgnore]
        public List<Element> WarningsElementColl { get; } = new List<Element>();

        [JsonIgnore]
        public List<Element> TypeParameterElementColl { get; } = new List<Element>();

        [JsonIgnore]
        public ObservableCollection<string> SystemSumParameters { get; private set; } = new ObservableCollection<string>();

        [JsonIgnore]
        public ObservableCollection<OVVK_SystemManagerRowM> SystemRows { get; private set; } = new ObservableCollection<OVVK_SystemManagerRowM>();

        [JsonIgnore]
        public bool HasParameterErrors { get; private set; }

        [JsonIgnore]
        public bool IsClosed { get; internal set; }

        public string ParameterName
        {
            get => _parameterName;
            set
            {
                if (_parameterName != (value ?? string.Empty))
                    SelectedSystemRow = null;

                _parameterName = value ?? string.Empty;
                NotifyPropertyChanged();
                NotifyPropertyChanged(nameof(ParameterDisplayName));
                NotifyState();
            }
        }

        public string SysNameSeparator
        {
            get => _sysNameSeparator;
            set => _sysNameSeparator = string.IsNullOrEmpty(value) ? "/" : value;
        }

        [JsonIgnore]
        public OVVK_SystemManagerRowM SelectedSystemRow
        {
            get => SelectedSystemRows.FirstOrDefault();
            set
            {
                SetSelectedRows(value == null ? new OVVK_SystemManagerRowM[0] : new[] { value });
            }
        }

        internal void SetSelectedRows(IEnumerable<OVVK_SystemManagerRowM> selectedRows)
        {
            var selected = new HashSet<OVVK_SystemManagerRowM>(selectedRows);
            foreach (OVVK_SystemManagerRowM row in SystemRows)
                row.IsSelected = selected.Contains(row);
            NotifySelection();
        }

        [JsonIgnore]
        public OVVK_SystemManagerRowM[] SelectedSystemRows => SystemRows.Where(row => row.IsSelected).ToArray();

        [JsonIgnore]
        public string[] SelectedSystemNames => SelectedSystemRows
            .Where(row => row.Kind == OVVK_SystemManagerRowKind.System)
            .Select(row => row.SystemName).Distinct().ToArray();

        [JsonIgnore]
        public bool HasSelectedErrors => SelectedSystemRows.Any(row => row.IsError);

        [JsonIgnore]
        public bool HasSelectedWarnings => SelectedSystemRows.Any(row => row.IsWarning);

        [JsonIgnore]
        public string CreateViewsButtonText => $"Создать 3D-виды ({SelectedSystemNames.Length})";

        [JsonIgnore]
        public string SelectionSummary => $"Выбрано строк: {SelectedSystemRows.Length}. Систем для 3D-видов: {SelectedSystemNames.Length}.";

        [JsonIgnore]
        public bool IsBusy
        {
            get => _isBusy;
            internal set
            {
                _isBusy = value;
                NotifyPropertyChanged();
                NotifyState();
            }
        }

        [JsonIgnore]
        public bool IsDocumentActive
        {
            get => _isDocumentActive;
            internal set
            {
                if (_isDocumentActive == value)
                    return;

                _isDocumentActive = value;
                NotifyPropertyChanged();
                NotifyState();
            }
        }

        [JsonIgnore]
        public string ParameterDisplayName => string.IsNullOrWhiteSpace(ParameterName) ? "Параметр не выбран" : ParameterName;

        [JsonIgnore]
        public bool CanRefresh => !IsClosed && IsDocumentActive && !IsBusy;

        [JsonIgnore]
        public bool CanSelectParameter => CanRefresh && ElementColl.Count > 0;

        [JsonIgnore]
        public bool CanFillSystemName => CanRefresh && !string.IsNullOrWhiteSpace(ParameterName);

        [JsonIgnore]
        public bool CanRun => CanSelectParameter && !HasParameterErrors && !HasTypeParameters && !string.IsNullOrWhiteSpace(ParameterName);

        [JsonIgnore]
        public bool CanMergeSystems => CanRun && SystemSumParameters.Count > 1;

        [JsonIgnore]
        public bool CanCreateViews => CanRefresh && !string.IsNullOrWhiteSpace(ParameterName) && SelectedSystemNames.Length > 0;

        [JsonIgnore]
        public bool CanSelectSystem => CanRefresh && SelectedSystemRows.Length > 0;

        [JsonIgnore]
        public string SelectSystemButtonText => SelectedSystemRows.Any(row => row.IsIssue)
            ? "Выделить элементы"
            : SelectedSystemRows.Length > 1 ? "Выделить системы" : "Выделить систему";

        [JsonIgnore]
        public bool HasTypeParameters => TypeParameterElementColl.Count > 0;

        [JsonIgnore]
        public string TypeParameterWarning => HasTypeParameters
            ? $"ВАЖНО: У {TypeParameterElementColl.Count} элементов «{ParameterName}» — параметр типа. "
                + "Заполнение меняет значение у всех экземпляров типа. При повторной записи остаётся последнее значение. "
                + "Для разных систем используй параметр экземпляра. Выбери оранжевую строку, чтобы выделить эти элементы."
            : string.Empty;

        [JsonIgnore]
        public string UserMainStatus => $"Элементов: {ElementColl.Count}   Систем: {SystemSumParameters.Count}   С замечаниями: {WarningsElementColl.Count}"
            + (HasTypeParameters ? $"   Параметр типа: {TypeParameterElementColl.Count}" : string.Empty);

        [JsonIgnore]
        public string UserHelp
        {
            get
            {
                if (!IsDocumentActive)
                    return $"ВАЖНО: Переключись в модель «{DocumentTitle}», для которой открыто окно.";
                if (IsBusy)
                    return "Выполняется обработка. Дождись завершения.";
                if (!string.IsNullOrEmpty(_operationMessage))
                    return "ВАЖНО: " + _operationMessage;
                if (ElementColl.Count == 0)
                    return "ВАЖНО: В модели нет элементов ОВВК для анализа.";
                if (string.IsNullOrWhiteSpace(ParameterName))
                    return "ВАЖНО: Выбери текстовый параметр для группирования элементов по системам.";
                if (HasParameterErrors)
                    return "ВАЖНО: Выбери красную строку и нажми «Выделить элементы», чтобы проверить ошибки параметра системы.";

                return string.Empty;
            }
        }


        /// <summary>
        /// Чтение модели выполняется только из ExtCmd, очереди команд или обработчика Revit.
        /// </summary>
        public void Refresh()
        {
            _operationMessage = string.Empty;
            ElementColl = CollectElements(CurrentDoc).ToList();
            UpdateSystemParamData();
        }

        /// <summary>
        /// Единый набор категорий и исключений для списка менеджера и заполнения параметра.
        /// Если передан viewId, рассматриваются только элементы этого вида.
        /// </summary>
        internal static Element[] CollectElements(Document doc, ElementId viewId = null, Action<int, int> reportProgress = null)
        {
            if (viewId != null && !FilteredElementCollector.IsViewValidForElementIteration(doc, viewId))
                throw new InvalidOperationException("На текущем виде нельзя собрать элементы. Открой план или 3D-вид либо выбери всю модель.");

            var elements = new List<Element>();
            int processedCategories = 0;
            int categoryCount = _familyCategories.Length + _curveCategories.Length;

            foreach (BuiltInCategory category in _familyCategories)
            {
                reportProgress?.Invoke(processedCategories, categoryCount);
                using (var collector = viewId == null ? new FilteredElementCollector(doc) : new FilteredElementCollector(doc, viewId))
                {
                    elements.AddRange(collector
                        .OfClass(typeof(FamilyInstance))
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(instance => !_excludedFamilyPrefixes.Any(prefix => instance.Symbol.FamilyName.StartsWith(prefix))));
                }

                processedCategories++;
            }

            foreach (BuiltInCategory category in _curveCategories)
            {
                reportProgress?.Invoke(processedCategories, categoryCount);
                using (var collector = viewId == null ? new FilteredElementCollector(doc) : new FilteredElementCollector(doc, viewId))
                {
                    elements.AddRange(collector
                        .OfCategory(category)
                        .WhereElementIsNotElementType());
                }

                processedCategories++;
            }

            reportProgress?.Invoke(categoryCount, categoryCount);
            return elements.ToArray();
        }

        public void UpdateSystemParamData()
        {
            var selectedKeys = SelectedSystemRows.Select(row => new { row.Kind, row.SystemName }).ToArray();
            var systemNames = new HashSet<string>();
            var typeParameters = new Dictionary<ElementId, Parameter>();
            var invalidNames = new HashSet<string>();
            _systemElementIds.Clear();
            _errorElementIds.Clear();
            WarningsElementColl.Clear();
            TypeParameterElementColl.Clear();
            HasParameterErrors = false;

            if (!string.IsNullOrWhiteSpace(ParameterName))
            {
                foreach (Element element in ElementColl)
                {
                    Parameter parameter = element.LookupParameter(ParameterName);
                    bool isTypeParameter = false;
                    if (parameter == null)
                    {
                        // Как при заполнении: сначала экземпляр, затем его тип.
                        ElementId typeId = element.GetTypeId();
                        if (!typeParameters.TryGetValue(typeId, out Parameter typeParameter))
                        {
                            typeParameter = CurrentDoc.GetElement(typeId)?.LookupParameter(ParameterName);
                            typeParameters.Add(typeId, typeParameter);
                        }

                        if (typeParameter != null)
                        {
                            TypeParameterElementColl.Add(element);
                            parameter = typeParameter;
                            isTypeParameter = true;
                        }
                    }

                    if (parameter == null || parameter.StorageType != StorageType.String)
                    {
                        AddError(element, OVVK_SystemManagerRowKind.InvalidParameterError);
                        continue;
                    }

                    string systemName = parameter.AsString();

                    // Пустое собственное значение нельзя скрывать заполненным параметром родителя.
                    if (!parameter.HasValue || string.IsNullOrWhiteSpace(systemName))
                    {
                        AddError(element, OVVK_SystemManagerRowKind.EmptyParameterError);
                        continue;
                    }

                    if (SystemNameValidation.IsServiceValue(systemName))
                    {
                        AddError(element, OVVK_SystemManagerRowKind.SystemNameError);
                        invalidNames.Add(systemName);
                        continue;
                    }

                    // Заполненные параметры типа остаются предупреждением, а не отдельными системами.
                    if (isTypeParameter)
                        continue;

                    systemNames.Add(systemName);
                    if (!_systemElementIds.TryGetValue(systemName, out List<ElementId> elementIds))
                    {
                        elementIds = new List<ElementId>();
                        _systemElementIds.Add(systemName, elementIds);
                    }

                    elementIds.Add(element.Id);
                }
            }

            SystemSumParameters = new ObservableCollection<string>(systemNames.OrderBy(name => name));
            var rows = new List<OVVK_SystemManagerRowM>();
            foreach (var error in _errorElementIds.OrderBy(pair => pair.Key))
            {
                string details = error.Key == OVVK_SystemManagerRowKind.SystemNameError
                    ? "Значения: " + string.Join("; ", invalidNames.OrderBy(name => name).Take(10))
                    : null;
                rows.Add(new OVVK_SystemManagerRowM(error.Key, error.Value.Count, details));
            }
            if (HasTypeParameters)
                rows.Add(new OVVK_SystemManagerRowM(TypeParameterElementColl.Count));
            rows.AddRange(SystemSumParameters.Select(name => new OVVK_SystemManagerRowM(name)));

            foreach (OVVK_SystemManagerRowM oldRow in SystemRows)
                oldRow.PropertyChanged -= OnRowPropertyChanged;
            foreach (OVVK_SystemManagerRowM row in rows)
            {
                row.IsSelected = selectedKeys.Any(key => key.Kind == row.Kind && key.SystemName == row.SystemName);
                row.PropertyChanged += OnRowPropertyChanged;
            }
            SystemRows = new ObservableCollection<OVVK_SystemManagerRowM>(rows);
            NotifyPropertyChanged(nameof(ElementColl));
            NotifyPropertyChanged(nameof(WarningsElementColl));
            NotifyPropertyChanged(nameof(TypeParameterElementColl));
            NotifyPropertyChanged(nameof(SystemSumParameters));
            NotifyPropertyChanged(nameof(SystemRows));

            // Восстанавливаем выбор после уведомлений WPF о замене коллекции.
            SetSelectedRows(SystemRows.Where(row => selectedKeys.Any(key => key.Kind == row.Kind && key.SystemName == row.SystemName)));
            NotifyState();
        }

        internal ElementId[] GetSystemElementIds(OVVK_SystemManagerRowM row)
        {
            if (row?.IsError == true)
                return _errorElementIds.TryGetValue(row.Kind, out List<ElementId> errorIds)
                    ? errorIds.ToArray()
                    : new ElementId[0];

            if (row?.IsTypeParameterWarning == true)
                return TypeParameterElementColl.Select(element => element.Id).ToArray();

            string systemName = row?.SystemName;
            return systemName != null && _systemElementIds.TryGetValue(systemName, out List<ElementId> elementIds)
                ? elementIds.ToArray()
                : new ElementId[0];
        }

        internal void SetOperationMessage(string message)
        {
            _operationMessage = message ?? string.Empty;
            NotifyPropertyChanged(nameof(UserHelp));
        }

        public object ToJson()
        {
            return new { DocumentKey, ParameterName, SysNameSeparator };
        }


        private void OnRowPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(OVVK_SystemManagerRowM.IsSelected))
                NotifySelection();
        }

        private void NotifySelection()
        {
            NotifyPropertyChanged(nameof(SelectedSystemRow));
            NotifyPropertyChanged(nameof(SelectedSystemRows));
            NotifyPropertyChanged(nameof(SelectedSystemNames));
            NotifyPropertyChanged(nameof(HasSelectedErrors));
            NotifyPropertyChanged(nameof(HasSelectedWarnings));
            NotifyPropertyChanged(nameof(SelectSystemButtonText));
            NotifyPropertyChanged(nameof(CreateViewsButtonText));
            NotifyPropertyChanged(nameof(SelectionSummary));
            NotifyPropertyChanged(nameof(CanSelectSystem));
            NotifyPropertyChanged(nameof(CanCreateViews));
        }

        private void NotifyState()
        {
            NotifyPropertyChanged(nameof(CanRefresh));
            NotifyPropertyChanged(nameof(CanSelectParameter));
            NotifyPropertyChanged(nameof(CanFillSystemName));
            NotifyPropertyChanged(nameof(CanRun));
            NotifyPropertyChanged(nameof(CanMergeSystems));
            NotifyPropertyChanged(nameof(CanCreateViews));
            NotifyPropertyChanged(nameof(CanSelectSystem));
            NotifyPropertyChanged(nameof(HasTypeParameters));
            NotifyPropertyChanged(nameof(TypeParameterWarning));
            NotifyPropertyChanged(nameof(UserMainStatus));
            NotifyPropertyChanged(nameof(UserHelp));
        }

        private void AddError(Element element, OVVK_SystemManagerRowKind kind)
        {
            HasParameterErrors = true;
            WarningsElementColl.Add(element);
            if (!_errorElementIds.TryGetValue(kind, out List<ElementId> ids))
            {
                ids = new List<ElementId>();
                _errorElementIds.Add(kind, ids);
            }

            ids.Add(element.Id);
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
