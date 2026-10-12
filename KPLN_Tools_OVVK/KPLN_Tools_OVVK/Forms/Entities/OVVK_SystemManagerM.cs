using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker.Core;
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
        public ObservableCollection<string> SystemSumParameters { get; private set; } = new ObservableCollection<string>();

        [JsonIgnore]
        public bool HasParameterErrors { get; private set; }

        [JsonIgnore]
        public bool IsClosed { get; internal set; }

        public string ParameterName
        {
            get => _parameterName;
            set
            {
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
        public bool CanRun => CanSelectParameter && !HasParameterErrors && !string.IsNullOrWhiteSpace(ParameterName);

        [JsonIgnore]
        public bool CanMergeSystems => CanRun && SystemSumParameters.Count > 1;

        [JsonIgnore]
        public bool CanCreateViews => CanRun && SystemSumParameters.Count > 0;

        [JsonIgnore]
        public bool CanSelectWarnings => CanRefresh && WarningsElementColl.Count > 0;

        [JsonIgnore]
        public string UserMainStatus => $"Элементов: {ElementColl.Count}   Систем: {SystemSumParameters.Count}   С замечаниями: {WarningsElementColl.Count}";

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
                    return "ВАЖНО: У части элементов выбранный параметр отсутствует или имеет другой тип. Выдели элементы с ошибками.";
                if (WarningsElementColl.Count > 0)
                    return "ВАЖНО: Есть элементы с пустым значением системы. Проверь их перед обработкой.";

                return string.Empty;
            }
        }


        /// <summary>
        /// Чтение модели выполняется только из ExtCmd, очереди команд или обработчика Revit.
        /// </summary>
        public void Refresh()
        {
            _operationMessage = string.Empty;
            var elements = new List<Element>();

            foreach (BuiltInCategory category in _familyCategories)
            {
                elements.AddRange(new FilteredElementCollector(CurrentDoc)
                    .OfClass(typeof(FamilyInstance))
                    .OfCategory(category)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .Where(instance => !_excludedFamilyPrefixes.Any(prefix => instance.Symbol.FamilyName.StartsWith(prefix))));
            }

            foreach (BuiltInCategory category in _curveCategories)
            {
                elements.AddRange(new FilteredElementCollector(CurrentDoc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType());
            }

            ElementColl = elements;
            UpdateSystemParamData();
        }

        public void UpdateSystemParamData()
        {
            var systemNames = new HashSet<string>();
            WarningsElementColl.Clear();
            HasParameterErrors = false;

            if (!string.IsNullOrWhiteSpace(ParameterName))
            {
                foreach (Element element in ElementColl)
                {
                    Parameter parameter = element.LookupParameter(ParameterName);
                    if (parameter == null || parameter.StorageType != StorageType.String)
                    {
                        HasParameterErrors = true;
                        WarningsElementColl.Add(element);
                        continue;
                    }

                    string systemName = parameter.AsString();

                    // У вложенных общих семейств сохраняем поиск системы у главного родителя.
                    if (string.IsNullOrWhiteSpace(systemName) && element is FamilyInstance instance)
                    {
                        Element parent = instance;
                        while (parent is FamilyInstance family && family.SuperComponent != null)
                            parent = family.SuperComponent;

                        if (!parent.Id.Equals(element.Id))
                        {
                            Parameter parentParameter = parent.LookupParameter(ParameterName);
                            if (parentParameter != null && parentParameter.StorageType == StorageType.String)
                                systemName = parentParameter.AsString();
                        }
                    }

                    if (string.IsNullOrWhiteSpace(systemName))
                        WarningsElementColl.Add(element);
                    else
                        systemNames.Add(systemName);
                }
            }

            SystemSumParameters = new ObservableCollection<string>(systemNames.OrderBy(name => name));
            NotifyPropertyChanged(nameof(ElementColl));
            NotifyPropertyChanged(nameof(WarningsElementColl));
            NotifyPropertyChanged(nameof(SystemSumParameters));
            NotifyState();
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


        private void NotifyState()
        {
            NotifyPropertyChanged(nameof(CanRefresh));
            NotifyPropertyChanged(nameof(CanSelectParameter));
            NotifyPropertyChanged(nameof(CanRun));
            NotifyPropertyChanged(nameof(CanMergeSystems));
            NotifyPropertyChanged(nameof(CanCreateViews));
            NotifyPropertyChanged(nameof(CanSelectWarnings));
            NotifyPropertyChanged(nameof(UserMainStatus));
            NotifyPropertyChanged(nameof(UserHelp));
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
