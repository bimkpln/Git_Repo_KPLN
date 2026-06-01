using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using TextBox = System.Windows.Controls.TextBox;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class CollectionParametersWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;
        private readonly DeleteParametersExternalEventHandler _deleteParametersHandler;
        private readonly ExternalEvent _deleteParametersEvent;

        private string _windowTitle;
        private string _description;
        private string _status;
        private bool _isCompactMode;
        private ParameterUsageRow _deleteSelectionAnchorRow;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<ParameterUsageRow> Rows { get; private set; }

        public string WindowTitle
        {
            get { return _windowTitle; }
            private set
            {
                _windowTitle = value;
                OnPropertyChanged("WindowTitle");
            }
        }

        public string Description
        {
            get { return _description; }
            private set
            {
                _description = value;
                OnPropertyChanged("Description");
            }
        }

        public string Status
        {
            get { return _status; }
            private set
            {
                _status = value;
                OnPropertyChanged("Status");
            }
        }

        public bool IsCompactMode
        {
            get { return _isCompactMode; }
            set
            {
                if (_isCompactMode == value)
                {
                    return;
                }

                _isCompactMode = value;
                OnPropertyChanged("IsCompactMode");
                UpdateCompactModeDisplay();
            }
        }

        public CollectionParametersWindow(ExternalCommandData commandData)
        {
            _doc = commandData.Application.ActiveUIDocument.Document;
            _deleteParametersHandler = new DeleteParametersExternalEventHandler(this);
            _deleteParametersEvent = ExternalEvent.Create(_deleteParametersHandler);
            Rows = new ObservableCollection<ParameterUsageRow>();

            InitializeComponent();
            DataContext = this;

            ConfigureMode();
            LoadData();
        }

        private void ConfigureMode()
        {
            if (_doc.IsFamilyDocument)
            {
                WindowTitle = "KPLN. Параметры семейства";
                Description = "Анализ общих параметров открытого семейства: использование в формулах и размерных метках.";

                FamiliesColumn.Visibility = System.Windows.Visibility.Collapsed;
                ViewFiltersColumn.Visibility = System.Windows.Visibility.Collapsed;
                SchedulesColumn.Visibility = System.Windows.Visibility.Collapsed;
                ScheduleFiltersColumn.Visibility = System.Windows.Visibility.Collapsed;
                FormulaParametersColumn.Visibility = System.Windows.Visibility.Visible;
                DimensionsColumn.Visibility = System.Windows.Visibility.Visible;
            }
            else
            {
                WindowTitle = "KPLN. Параметры проекта";
                Description = "Анализ общих параметров проекта: категории из привязок параметров, фильтры видов и фильтры спецификаций.";

                FamiliesColumn.Visibility = System.Windows.Visibility.Visible;
                ViewFiltersColumn.Visibility = System.Windows.Visibility.Visible;
                SchedulesColumn.Visibility = System.Windows.Visibility.Visible;
                ScheduleFiltersColumn.Visibility = System.Windows.Visibility.Visible;
                FormulaParametersColumn.Visibility = System.Windows.Visibility.Collapsed;
                DimensionsColumn.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private void LoadData()
        {
            try
            {
                Rows.Clear();
                Status = "Идет сбор данных...";

                if (_doc.IsFamilyDocument)
                {
                    LoadFamilyParametersUsage();
                }
                else
                {
                    LoadProjectParametersUsage();
                }

                UpdateUsageColumnsVisibility();
                UpdateCompactModeDisplay();
                Status = string.Format("Найдено параметров: {0}", Rows.Count);
            }
            catch (Exception ex)
            {
                Status = "Ошибка сбора данных";
                TaskDialog.Show("Параметры", ex.Message);
            }
        }

        private void UpdateUsageColumnsVisibility()
        {
            if (_doc.IsFamilyDocument)
            {
                FamiliesColumn.Visibility = System.Windows.Visibility.Collapsed;
                ViewFiltersColumn.Visibility = System.Windows.Visibility.Collapsed;
                SchedulesColumn.Visibility = System.Windows.Visibility.Collapsed;
                ScheduleFiltersColumn.Visibility = System.Windows.Visibility.Collapsed;
                FormulaParametersColumn.Visibility = HasColumnData("FormulaParameters")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
                DimensionsColumn.Visibility = HasColumnData("Dimensions")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
            }
            else
            {
                FamiliesColumn.Visibility = HasColumnData("Families")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
                ViewFiltersColumn.Visibility = HasColumnData("ViewFilters")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
                SchedulesColumn.Visibility = HasColumnData("Schedules")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
                ScheduleFiltersColumn.Visibility = HasColumnData("ScheduleFilters")
                    ? System.Windows.Visibility.Visible
                    : System.Windows.Visibility.Collapsed;
                FormulaParametersColumn.Visibility = System.Windows.Visibility.Collapsed;
                DimensionsColumn.Visibility = System.Windows.Visibility.Collapsed;
            }
        }

        private bool HasColumnData(string propertyName)
        {
            foreach (ParameterUsageRow row in Rows)
            {
                if (!string.IsNullOrWhiteSpace(GetColumnValue(row, propertyName)))
                {
                    return true;
                }
            }

            return false;
        }

        private void UpdateCompactModeDisplay()
        {
            if (Rows == null)
            {
                return;
            }

            foreach (ParameterUsageRow row in Rows)
            {
                if (IsCompactMode)
                {
                    row.CollapseCompactCells();
                }

                row.UpdateCompactDisplay(IsCompactMode);
            }
        }

        private void LoadProjectParametersUsage()
        {
            List<ParameterUsageRow> parameterRows = GetSharedProjectParameters();
            Dictionary<long, ParameterUsageRow> rowsByParameterId = parameterRows
                .Where(x => IsValidElementId(x.ParameterId))
                .GroupBy(x => IDHelper.ElIdValue(x.ParameterId))
                .ToDictionary(x => x.Key, x => x.First());

            FillViewFiltersUsage(rowsByParameterId);
            FillScheduleFieldsUsage(rowsByParameterId);
            FillScheduleFiltersUsage(rowsByParameterId);

            foreach (ParameterUsageRow row in parameterRows.OrderBy(x => x.Name))
            {
                row.Families = JoinLines(row.FamiliesSet);
                row.ViewFilters = JoinLines(row.ViewFiltersSet);
                row.Schedules = JoinLines(row.SchedulesSet);
                row.ScheduleFilters = JoinLines(row.ScheduleFiltersSet);
                Rows.Add(row);
            }
        }

        private List<ParameterUsageRow> GetSharedProjectParameters()
        {
            Dictionary<string, ParameterUsageRow> rowsByGuid =
                new Dictionary<string, ParameterUsageRow>(StringComparer.OrdinalIgnoreCase);

            BindingMap bindingMap = _doc.ParameterBindings;
            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();

            while (iterator.MoveNext())
            {
                Definition definition = iterator.Key;
                Binding binding = iterator.Current as Binding;
                if (definition == null)
                {
                    continue;
                }

                ElementId parameterId = GetDefinitionElementId(definition);
                Guid guid;
                if (!TryGetSharedParameterGuid(definition, parameterId, out guid))
                {
                    continue;
                }

                if (!IsValidElementId(parameterId))
                {
                    parameterId = GetSharedParameterElementId(guid, definition.Name);
                }

                ParameterUsageRow row = GetOrCreateProjectParameterRow(rowsByGuid, guid, parameterId, definition.Name);
                row.IsProjectParameter = true;
                UpdateRowParameterId(row, parameterId);
                AddBindingCategoriesUsage(row, binding);
            }

            foreach (SharedParameterElement sharedParameter in GetSharedParameterElements())
            {
                if (sharedParameter == null || sharedParameter.GuidValue == Guid.Empty)
                {
                    continue;
                }

                GetOrCreateProjectParameterRow(
                    rowsByGuid,
                    sharedParameter.GuidValue,
                    sharedParameter.Id,
                    GetSharedParameterName(sharedParameter));
            }

            return rowsByGuid.Values.ToList();
        }

        private ParameterUsageRow GetOrCreateProjectParameterRow(
            Dictionary<string, ParameterUsageRow> rowsByGuid,
            Guid guid,
            ElementId parameterId,
            string parameterName)
        {
            string guidText = guid.ToString();

            ParameterUsageRow row;
            if (rowsByGuid.TryGetValue(guidText, out row))
            {
                if (string.IsNullOrWhiteSpace(row.Name))
                {
                    row.Name = parameterName ?? string.Empty;
                }

                UpdateRowParameterId(row, parameterId);
                return row;
            }

            row = new ParameterUsageRow
            {
                ParameterIdText = GetElementIdText(parameterId),
                Name = parameterName ?? string.Empty,
                Guid = guidText,
                ParameterGuid = guid,
                ParameterId = parameterId
            };

            rowsByGuid[guidText] = row;
            return row;
        }

        private void UpdateRowParameterId(ParameterUsageRow row, ElementId parameterId)
        {
            if (row == null || !IsValidElementId(parameterId) || IsValidElementId(row.ParameterId))
            {
                return;
            }

            row.ParameterId = parameterId;
            row.ParameterIdText = GetElementIdText(parameterId);
        }

        private IEnumerable<SharedParameterElement> GetSharedParameterElements()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(SharedParameterElement))
                .Cast<SharedParameterElement>();
        }

        private string GetSharedParameterName(SharedParameterElement sharedParameter)
        {
            if (sharedParameter == null)
            {
                return string.Empty;
            }

            try
            {
                Definition definition = sharedParameter.GetDefinition();
                if (definition != null && !string.IsNullOrWhiteSpace(definition.Name))
                {
                    return definition.Name;
                }
            }
            catch
            {
            }

            return sharedParameter.Name ?? string.Empty;
        }

        private void AddBindingCategoriesUsage(ParameterUsageRow row, Binding binding)
        {
            ElementBinding elementBinding = binding as ElementBinding;
            if (elementBinding == null || elementBinding.Categories == null)
            {
                return;
            }

            foreach (Category category in elementBinding.Categories)
            {
                if (category == null)
                {
                    continue;
                }

                row.FamiliesSet.Add(GetCategoryReferenceText(category));
            }
        }

        private string GetCategoryReferenceText(Category category)
        {
            return string.Format("{0} - {1}", IDHelper.ElIdValue(category.Id), category.Name);
        }

        private ElementId GetSharedParameterElementId(Guid guid, string parameterName)
        {
            SharedParameterElement sharedParameter = GetSharedParameterElements()
                .FirstOrDefault(x => x.GuidValue == guid ||
                                     GetSharedParameterName(x) == parameterName);

            return sharedParameter != null ? sharedParameter.Id : ElementId.InvalidElementId;
        }

        private ElementId GetDefinitionElementId(Definition definition)
        {
            InternalDefinition internalDefinition = definition as InternalDefinition;
            if (internalDefinition != null)
            {
                return internalDefinition.Id;
            }

            return ElementId.InvalidElementId;
        }

        private bool TryGetSharedParameterGuid(Definition definition, ElementId parameterId, out Guid guid)
        {
            ExternalDefinition externalDefinition = definition as ExternalDefinition;
            if (externalDefinition != null)
            {
                guid = externalDefinition.GUID;
                return guid != Guid.Empty;
            }

            if (IsValidElementId(parameterId))
            {
                SharedParameterElement sharedParameter = _doc.GetElement(parameterId) as SharedParameterElement;
                if (sharedParameter != null)
                {
                    guid = sharedParameter.GuidValue;
                    return guid != Guid.Empty;
                }
            }

            SharedParameterElement byName = GetSharedParameterElements()
                .FirstOrDefault(x => GetSharedParameterName(x) == definition.Name);

            if (byName != null)
            {
                guid = byName.GuidValue;
                return guid != Guid.Empty;
            }

            guid = Guid.Empty;
            return false;
        }

        private string GetNamedElementReferenceText(Element element)
        {
            if (element == null)
            {
                return string.Empty;
            }

            return string.Format("{0} - {1}", IDHelper.ElIdValue(element.Id), element.Name);
        }

        private string GetElementIdText(ElementId elementId)
        {
            return IsValidElementId(elementId) ? IDHelper.ElIdValue(elementId).ToString() : string.Empty;
        }

        private void FillViewFiltersUsage(Dictionary<long, ParameterUsageRow> rowsByParameterId)
        {
            if (rowsByParameterId.Count == 0)
            {
                return;
            }

            IEnumerable<ParameterFilterElement> filters = new FilteredElementCollector(_doc)
                .OfClass(typeof(ParameterFilterElement))
                .Cast<ParameterFilterElement>();

            foreach (ParameterFilterElement filter in filters)
            {
                HashSet<long> parameterIds = GetParameterIdsFromViewFilter(filter);
                foreach (long parameterId in parameterIds)
                {
                    ParameterUsageRow row;
                    if (rowsByParameterId.TryGetValue(parameterId, out row))
                    {
                        row.ViewFiltersSet.Add(GetNamedElementReferenceText(filter));
                    }
                }
            }
        }

        private HashSet<long> GetParameterIdsFromViewFilter(ParameterFilterElement filter)
        {
            HashSet<long> result = new HashSet<long>();

            AddRuleParameterIds(filter, result);

            object elementFilter = InvokeParameterlessMethod(filter, "GetElementFilter");
            AddElementFilterParameterIds(elementFilter, result);

            return result;
        }

        private void AddElementFilterParameterIds(object elementFilter, HashSet<long> result)
        {
            if (elementFilter == null)
            {
                return;
            }

            AddRuleParameterIds(elementFilter, result);

            object nestedFilters = InvokeParameterlessMethod(elementFilter, "GetFilters");
            AddNestedElementFilters(nestedFilters, result);

            object filtersProperty = GetPropertyValue(elementFilter, "Filters");
            AddNestedElementFilters(filtersProperty, result);
        }

        private void AddNestedElementFilters(object filtersObject, HashSet<long> result)
        {
            IEnumerable filters = filtersObject as IEnumerable;
            if (filters == null)
            {
                return;
            }

            foreach (object filter in filters)
            {
                AddElementFilterParameterIds(filter, result);
            }
        }

        private void AddRuleParameterIds(object source, HashSet<long> result)
        {
            object rulesObject = InvokeParameterlessMethod(source, "GetRules");
            IEnumerable rules = rulesObject as IEnumerable;
            if (rules == null)
            {
                return;
            }

            foreach (object rule in rules)
            {
                ElementId parameterId = GetRuleParameterId(rule);
                if (IsValidElementId(parameterId))
                {
                    result.Add(IDHelper.ElIdValue(parameterId));
                }
            }
        }

        private ElementId GetRuleParameterId(object rule)
        {
            if (rule == null)
            {
                return ElementId.InvalidElementId;
            }

            object directParameterId = InvokeParameterlessMethod(rule, "GetRuleParameter");
            ElementId parameterId = directParameterId as ElementId;
            if (IsValidElementId(parameterId))
            {
                return parameterId;
            }

            object ruleParameterProperty = GetPropertyValue(rule, "RuleParameter");
            parameterId = ruleParameterProperty as ElementId;
            if (IsValidElementId(parameterId))
            {
                return parameterId;
            }

            object valueProvider = InvokeParameterlessMethod(rule, "GetValueProvider") ?? GetPropertyValue(rule, "ValueProvider");
            if (valueProvider != null)
            {
                object providerParameter = GetPropertyValue(valueProvider, "ParameterId") ??
                                           GetPropertyValue(valueProvider, "Parameter") ??
                                           InvokeParameterlessMethod(valueProvider, "GetParameter");

                parameterId = providerParameter as ElementId;
                if (IsValidElementId(parameterId))
                {
                    return parameterId;
                }
            }

            return ElementId.InvalidElementId;
        }

        private object InvokeParameterlessMethod(object source, string methodName)
        {
            if (source == null)
            {
                return null;
            }

            MethodInfo method = source.GetType().GetMethod(methodName, BindingFlags.Instance | BindingFlags.Public);
            if (method == null)
            {
                return null;
            }

            try
            {
                return method.Invoke(source, null);
            }
            catch
            {
                return null;
            }
        }

        private object InvokeMethod(object source, string methodName, params object[] args)
        {
            if (source == null)
            {
                return null;
            }

            MethodInfo[] methods = source.GetType()
                .GetMethods(BindingFlags.Instance | BindingFlags.Public)
                .Where(x => x.Name == methodName && x.GetParameters().Length == args.Length)
                .ToArray();

            foreach (MethodInfo method in methods)
            {
                try
                {
                    return method.Invoke(source, args);
                }
                catch
                {
                }
            }

            return null;
        }

        private object GetPropertyValue(object source, string propertyName)
        {
            if (source == null)
            {
                return null;
            }

            PropertyInfo property = source.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
            if (property == null)
            {
                return null;
            }

            try
            {
                return property.GetValue(source, null);
            }
            catch
            {
                return null;
            }
        }

        private void FillScheduleFiltersUsage(Dictionary<long, ParameterUsageRow> rowsByParameterId)
        {
            if (rowsByParameterId.Count == 0)
            {
                return;
            }

            IEnumerable<ViewSchedule> schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>();

            foreach (ViewSchedule schedule in schedules)
            {
                if (schedule.IsTemplate)
                {
                    continue;
                }

                try
                {
                    ScheduleDefinition definition = schedule.Definition;
                    int filterCount = definition.GetFilterCount();

                    for (int i = 0; i < filterCount; i++)
                    {
                        ScheduleFilter filter = definition.GetFilter(i);
                        ScheduleField field = definition.GetField(filter.FieldId);
                        ElementId parameterId = field.ParameterId;

                        if (!IsValidElementId(parameterId))
                        {
                            continue;
                        }

                        ParameterUsageRow row;
                        if (rowsByParameterId.TryGetValue(IDHelper.ElIdValue(parameterId), out row))
                        {
                            row.ScheduleFiltersSet.Add(GetNamedElementReferenceText(schedule));
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }
        }

        private void FillScheduleFieldsUsage(Dictionary<long, ParameterUsageRow> rowsByParameterId)
        {
            if (rowsByParameterId.Count == 0)
            {
                return;
            }

            IEnumerable<ViewSchedule> schedules = new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>();

            foreach (ViewSchedule schedule in schedules)
            {
                if (schedule.IsTemplate)
                {
                    continue;
                }

                try
                {
                    HashSet<long> parameterIds = GetScheduleFieldParameterIds(schedule.Definition);
                    foreach (long parameterId in parameterIds)
                    {
                        ParameterUsageRow row;
                        if (rowsByParameterId.TryGetValue(parameterId, out row))
                        {
                            row.SchedulesSet.Add(GetNamedElementReferenceText(schedule));
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }
        }

        private HashSet<long> GetScheduleFieldParameterIds(ScheduleDefinition definition)
        {
            HashSet<long> result = new HashSet<long>();
            if (definition == null)
            {
                return result;
            }

            IEnumerable fieldOrder = InvokeParameterlessMethod(definition, "GetFieldOrder") as IEnumerable;
            if (fieldOrder != null)
            {
                foreach (object fieldId in fieldOrder)
                {
                    AddScheduleFieldParameterId(definition, fieldId, result);
                }

                return result;
            }

            object fieldCountObject = InvokeParameterlessMethod(definition, "GetFieldCount");
            if (!(fieldCountObject is int))
            {
                return result;
            }

            int fieldCount = (int)fieldCountObject;
            for (int i = 0; i < fieldCount; i++)
            {
                object fieldId = InvokeMethod(definition, "GetFieldId", i);
                AddScheduleFieldParameterId(definition, fieldId ?? (object)i, result);
            }

            return result;
        }

        private void AddScheduleFieldParameterId(ScheduleDefinition definition, object fieldKey, HashSet<long> result)
        {
            object field = InvokeMethod(definition, "GetField", fieldKey);
            ElementId parameterId = GetPropertyValue(field, "ParameterId") as ElementId;
            if (IsValidElementId(parameterId))
            {
                result.Add(IDHelper.ElIdValue(parameterId));
            }
        }

        private void LoadFamilyParametersUsage()
        {
            FamilyManager familyManager = _doc.FamilyManager;
            List<FamilyParameter> allParameters = familyManager.Parameters
                .Cast<FamilyParameter>()
                .ToList();

            List<FamilyParameter> sharedParameters = allParameters
                .Where(x => x.IsShared)
                .OrderBy(x => GetFamilyParameterName(x))
                .ToList();

            Dictionary<string, SortedSet<string>> formulasByGuid = GetFamilyFormulaUsage(allParameters, sharedParameters);
            Dictionary<string, SortedSet<string>> dimensionsByGuid = GetFamilyDimensionUsage(sharedParameters);

            foreach (FamilyParameter parameter in sharedParameters)
            {
                string guidText = parameter.GUID.ToString();

                ParameterUsageRow row = new ParameterUsageRow
                {
                    ParameterIdText = GetElementIdText(parameter.Id),
                    Name = GetFamilyParameterName(parameter),
                    Guid = guidText,
                    IsProjectParameter = true,
                    ParameterGuid = parameter.GUID,
                    ParameterId = parameter.Id,
                    FormulaParameters = JoinLines(GetSetValue(formulasByGuid, guidText)),
                    Dimensions = JoinLines(GetSetValue(dimensionsByGuid, guidText))
                };

                Rows.Add(row);
            }
        }

        private Dictionary<string, SortedSet<string>> GetFamilyFormulaUsage(
            List<FamilyParameter> allParameters,
            List<FamilyParameter> sharedParameters)
        {
            Dictionary<string, SortedSet<string>> result = CreateGuidStringDictionary(sharedParameters);

            foreach (FamilyParameter sourceParameter in sharedParameters)
            {
                string parameterName = GetFamilyParameterName(sourceParameter);
                string guidText = sourceParameter.GUID.ToString();

                foreach (FamilyParameter formulaParameter in allParameters)
                {
                    string formula = GetFamilyParameterFormula(formulaParameter);
                    if (formulaParameter == sourceParameter || string.IsNullOrWhiteSpace(formula))
                    {
                        continue;
                    }

                    if (FormulaContainsParameter(formula, parameterName))
                    {
                        result[guidText].Add(GetFamilyParameterName(formulaParameter));
                    }
                }
            }

            return result;
        }

        private string GetFamilyParameterName(FamilyParameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                return parameter.Definition != null ? parameter.Definition.Name : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetFamilyParameterFormula(FamilyParameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                return parameter.Formula ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private bool FormulaContainsParameter(string formula, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(parameterName))
            {
                return false;
            }

            string pattern = string.Format(
                @"(?<![\p{{L}}\p{{N}}_]){0}(?![\p{{L}}\p{{N}}_])",
                Regex.Escape(parameterName));

            return Regex.IsMatch(formula, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private Dictionary<string, SortedSet<string>> GetFamilyDimensionUsage(List<FamilyParameter> sharedParameters)
        {
            Dictionary<string, SortedSet<string>> result = CreateGuidStringDictionary(sharedParameters);
            Dictionary<long, string> guidByParameterId = sharedParameters
                .Where(x => IsValidElementId(x.Id))
                .GroupBy(x => IDHelper.ElIdValue(x.Id))
                .ToDictionary(x => x.Key, x => x.First().GUID.ToString());

            List<Dimension> dimensions = new FilteredElementCollector(_doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .ToList();

            Dictionary<long, SortedSet<string>> viewNamesByDimensionId = GetDimensionViewNames(dimensions);

            foreach (Dimension dimension in dimensions)
            {
                FamilyParameter label;
                try
                {
                    label = dimension.FamilyLabel;
                }
                catch
                {
                    continue;
                }

                if (label == null || !IsValidElementId(label.Id))
                {
                    continue;
                }

                string guidText;
                if (!guidByParameterId.TryGetValue(IDHelper.ElIdValue(label.Id), out guidText))
                {
                    continue;
                }

                result[guidText].Add(GetDimensionReferenceText(dimension, viewNamesByDimensionId));
            }

            return result;
        }

        private Dictionary<long, SortedSet<string>> GetDimensionViewNames(List<Dimension> dimensions)
        {
            Dictionary<long, SortedSet<string>> result = dimensions
                .Where(x => IsValidElementId(x.Id))
                .GroupBy(x => IDHelper.ElIdValue(x.Id))
                .ToDictionary(
                    x => x.Key,
                    x => new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase));

            foreach (Dimension dimension in dimensions)
            {
                View ownerView = GetOwnerView(dimension);
                SortedSet<string> viewNames;
                if (ownerView != null && result.TryGetValue(IDHelper.ElIdValue(dimension.Id), out viewNames))
                {
                    viewNames.Add(ownerView.Name);
                }
            }

            List<View> views = new FilteredElementCollector(_doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(x => !x.IsTemplate)
                .ToList();

            foreach (View view in views)
            {
                try
                {
                    IEnumerable<Dimension> viewDimensions = new FilteredElementCollector(_doc, view.Id)
                        .OfClass(typeof(Dimension))
                        .Cast<Dimension>();

                    foreach (Dimension dimension in viewDimensions)
                    {
                        SortedSet<string> viewNames;
                        if (result.TryGetValue(IDHelper.ElIdValue(dimension.Id), out viewNames))
                        {
                            viewNames.Add(view.Name);
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return result;
        }

        private string GetDimensionReferenceText(Dimension dimension, Dictionary<long, SortedSet<string>> viewNamesByDimensionId)
        {
            SortedSet<string> viewNames;
            long dimensionId = IDHelper.ElIdValue(dimension.Id);
            if (viewNamesByDimensionId.TryGetValue(dimensionId, out viewNames) && viewNames.Count > 0)
            {
                return string.Format(
                    "{0} ({1})",
                    dimensionId,
                    string.Join("; ", viewNames));
            }

            return string.Format("{0} (вид не найден)", dimensionId);
        }

        private View GetOwnerView(Element element)
        {
            if (!IsValidElementId(element.OwnerViewId))
            {
                return null;
            }

            return _doc.GetElement(element.OwnerViewId) as View;
        }

        private Dictionary<string, SortedSet<string>> CreateGuidStringDictionary(List<FamilyParameter> parameters)
        {
            Dictionary<string, SortedSet<string>> result = new Dictionary<string, SortedSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (FamilyParameter parameter in parameters)
            {
                result[parameter.GUID.ToString()] = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
            }

            return result;
        }

        private SortedSet<string> GetSetValue(Dictionary<string, SortedSet<string>> source, string key)
        {
            SortedSet<string> value;
            if (source.TryGetValue(key, out value))
            {
                return value;
            }

            return new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
        }

        private string JoinLines(IEnumerable<string> values)
        {
            return string.Join(Environment.NewLine, values.Where(x => !string.IsNullOrWhiteSpace(x)));
        }

        private bool IsValidElementId(ElementId elementId)
        {
            return elementId != null &&
                   elementId != ElementId.InvalidElementId &&
                   IDHelper.ElIdValue(elementId) > 0;
        }

        private void DeleteParameterButton_Click(object sender, RoutedEventArgs e)
        {
            List<ParameterUsageRow> selectedRows = GetCheckedRows();
            if (selectedRows.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Отметьте галками один или несколько параметров для удаления.",
                    "Удаление параметров",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            MessageBoxResult confirmation = MessageBox.Show(
                this,
                selectedRows.Count == 1
                    ? "Удалить отмеченный параметр?"
                    : string.Format("Удалить отмеченные параметры: {0}?", selectedRows.Count),
                "Удаление параметров",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning,
                MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            _deleteParametersHandler.SetRows(selectedRows);

            try
            {
                _deleteParametersEvent.Raise();
                Status = "Запрошено удаление параметров...";
            }
            catch (Exception ex)
            {
                Status = "Ошибка запуска удаления";
                MessageBox.Show(this, ex.Message, "Удаление параметров", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteParametersInRevitContext(List<ParameterUsageRow> selectedRows)
        {
            int deletedCount = 0;
            List<string> errors = new List<string>();

            foreach (ParameterUsageRow row in selectedRows)
            {
                Transaction transaction = null;
                try
                {
                    transaction = new Transaction(_doc, "KPLN. Удаление общего параметра");
                    transaction.Start();

                    if (_doc.IsFamilyDocument)
                    {
                        DeleteFamilyParameter(row);
                    }
                    else
                    {
                        DeleteProjectParameter(row);
                    }

                    TransactionStatus status = transaction.Commit();
                    if (status != TransactionStatus.Committed)
                    {
                        throw new InvalidOperationException(string.Format("Revit отменил транзакцию: {0}", status));
                    }

                    row.IsSelectedForDelete = false;
                    deletedCount++;
                }
                catch (Exception ex)
                {
                    TryRollBackTransaction(transaction);
                    errors.Add(string.Format("{0}: {1}", GetRowDisplayName(row), ex.Message));
                }
            }

            LoadData();
            Status = string.Format("Удалено параметров: {0}", deletedCount);

            if (errors.Count > 0)
            {
                TaskDialog.Show("Удаление параметров", BuildErrorMessage(errors));
            }
        }

        private void TryRollBackTransaction(Transaction transaction)
        {
            if (transaction == null)
            {
                return;
            }

            try
            {
                if (transaction.GetStatus() == TransactionStatus.Started)
                {
                    transaction.RollBack();
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Удаление параметров", ex.Message);
            }
        }

        private void DeleteProjectParameter(ParameterUsageRow row)
        {
            if (row == null || !IsValidElementId(row.ParameterId))
            {
                throw new InvalidOperationException("не найден ElementId параметра");
            }

            _doc.Delete(row.ParameterId);
        }

        private void DeleteFamilyParameter(ParameterUsageRow row)
        {
            FamilyParameter parameter = GetFamilyParameterByGuid(row != null ? row.ParameterGuid : Guid.Empty);
            if (parameter == null)
            {
                throw new InvalidOperationException("параметр не найден в семействе");
            }

            _doc.FamilyManager.RemoveParameter(parameter);
        }

        private FamilyParameter GetFamilyParameterByGuid(Guid guid)
        {
            if (guid == Guid.Empty || !_doc.IsFamilyDocument)
            {
                return null;
            }

            FamilyManager familyManager = _doc.FamilyManager;
            foreach (FamilyParameter parameter in familyManager.Parameters.Cast<FamilyParameter>())
            {
                if (parameter.IsShared && parameter.GUID == guid)
                {
                    return parameter;
                }
            }

            return null;
        }

        private List<ParameterUsageRow> GetCheckedRows()
        {
            return Rows
                .Where(x => x.IsSelectedForDelete)
                .OrderBy(x => Rows.IndexOf(x))
                .ToList();
        }

        private void CompactCellTextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!IsCompactMode)
            {
                return;
            }

            FrameworkElement element = sender as FrameworkElement;
            ParameterUsageRow row = element != null
                ? element.DataContext as ParameterUsageRow
                : null;
            string cellKey = element != null
                ? element.Tag as string
                : null;

            if (row == null || string.IsNullOrWhiteSpace(cellKey) || row.GetCompactItemCount(cellKey) <= 1)
            {
                return;
            }

            row.ToggleCompactCell(cellKey, IsCompactMode);
            e.Handled = true;
        }

        private void CompactCellTextBox_MouseEnter(object sender, MouseEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null && IsCompactTextBoxCollapsed(textBox))
            {
                textBox.Cursor = Cursors.Hand;
            }
        }

        private void CompactCellTextBox_MouseLeave(object sender, MouseEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (textBox != null)
            {
                textBox.ClearValue(CursorProperty);
            }
        }

        private bool IsCompactTextBoxCollapsed(FrameworkElement element)
        {
            if (!IsCompactMode)
            {
                return false;
            }

            ParameterUsageRow row = element != null
                ? element.DataContext as ParameterUsageRow
                : null;
            string cellKey = element != null
                ? element.Tag as string
                : null;

            return row != null &&
                   !string.IsNullOrWhiteSpace(cellKey) &&
                   row.IsCompactCellCollapsed(cellKey);
        }

        private void DeleteSelectionCell_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            FrameworkElement element = sender as FrameworkElement;
            ParameterUsageRow clickedRow = element != null
                ? element.DataContext as ParameterUsageRow
                : null;

            if (clickedRow == null)
            {
                return;
            }

            e.Handled = true;

            bool isShiftPressed = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            bool isSelected = !clickedRow.IsSelectedForDelete;
            if (!isShiftPressed)
            {
                ParametersGrid.SelectedCells.Clear();
                clickedRow.IsSelectedForDelete = isSelected;
                _deleteSelectionAnchorRow = clickedRow;
                return;
            }

            List<ParameterUsageRow> rowsToUpdate = GetDeleteSelectionRows(clickedRow, isShiftPressed);

            foreach (ParameterUsageRow row in rowsToUpdate)
            {
                row.IsSelectedForDelete = isSelected;
            }

            if (_deleteSelectionAnchorRow == null)
            {
                _deleteSelectionAnchorRow = clickedRow;
            }
        }

        private void CollectionParametersWindow_PreviewKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.LeftShift || e.Key == Key.RightShift)
            {
                _deleteSelectionAnchorRow = null;
            }
        }

        private List<ParameterUsageRow> GetDeleteSelectionRows(ParameterUsageRow clickedRow, bool isShiftPressed)
        {
            if (!isShiftPressed)
            {
                return new List<ParameterUsageRow> { clickedRow };
            }

            List<ParameterUsageRow> selectedRows = GetSelectedParameterRowsInVisualOrder();
            if (selectedRows.Count > 1 && selectedRows.Contains(clickedRow))
            {
                return selectedRows;
            }

            if (_deleteSelectionAnchorRow != null)
            {
                return GetVisualRowsBetween(_deleteSelectionAnchorRow, clickedRow);
            }

            return new List<ParameterUsageRow> { clickedRow };
        }

        private List<ParameterUsageRow> GetSelectedParameterRowsInVisualOrder()
        {
            HashSet<ParameterUsageRow> selectedRows = new HashSet<ParameterUsageRow>(
                ParametersGrid.SelectedCells
                    .Select(x => x.Item)
                    .OfType<ParameterUsageRow>());

            if (selectedRows.Count == 0)
            {
                return new List<ParameterUsageRow>();
            }

            return ParametersGrid.Items
                .Cast<object>()
                .OfType<ParameterUsageRow>()
                .Where(x => selectedRows.Contains(x))
                .ToList();
        }

        private List<ParameterUsageRow> GetVisualRowsBetween(ParameterUsageRow fromRow, ParameterUsageRow toRow)
        {
            List<ParameterUsageRow> visibleRows = ParametersGrid.Items
                .Cast<object>()
                .OfType<ParameterUsageRow>()
                .ToList();

            int fromIndex = visibleRows.IndexOf(fromRow);
            int toIndex = visibleRows.IndexOf(toRow);
            if (fromIndex < 0 || toIndex < 0)
            {
                return new List<ParameterUsageRow> { toRow };
            }

            int startIndex = Math.Min(fromIndex, toIndex);
            int endIndex = Math.Max(fromIndex, toIndex);
            List<ParameterUsageRow> rows = new List<ParameterUsageRow>();

            for (int i = startIndex; i <= endIndex; i++)
            {
                rows.Add(visibleRows[i]);
            }

            return rows;
        }

        private class DeleteParametersExternalEventHandler : IExternalEventHandler
        {
            private readonly CollectionParametersWindow _window;
            private List<ParameterUsageRow> _rows;

            internal DeleteParametersExternalEventHandler(CollectionParametersWindow window)
            {
                _window = window;
                _rows = new List<ParameterUsageRow>();
            }

            internal void SetRows(List<ParameterUsageRow> rows)
            {
                _rows = rows != null
                    ? rows.ToList()
                    : new List<ParameterUsageRow>();
            }

            public void Execute(UIApplication app)
            {
                List<ParameterUsageRow> rows = _rows.ToList();
                _rows.Clear();

                try
                {
                    _window.DeleteParametersInRevitContext(rows);
                }
                catch (Exception ex)
                {
                    _window.Status = "Ошибка удаления параметров";
                    TaskDialog.Show("Удаление параметров", ex.Message);
                }
            }

            public string GetName()
            {
                return "KPLN. Удаление общих параметров";
            }
        }

        private string GetRowDisplayName(ParameterUsageRow row)
        {
            if (row == null)
            {
                return "Параметр";
            }

            if (!string.IsNullOrWhiteSpace(row.Name))
            {
                return row.Name;
            }

            return !string.IsNullOrWhiteSpace(row.Guid) ? row.Guid : "Параметр";
        }

        private string BuildErrorMessage(List<string> errors)
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Часть параметров не удалось удалить:");

            foreach (string error in errors)
            {
                builder.AppendLine(error);
            }

            return builder.ToString();
        }

        private void ExportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog
            {
                AddExtension = true,
                DefaultExt = ".xlsx",
                FileName = GetDefaultReportFileName(),
                Filter = "Книга Excel (*.xlsx)|*.xlsx",
                OverwritePrompt = true,
                Title = "Сохранить отчет"
            };

            bool? result = dialog.ShowDialog(this);
            if (result != true)
            {
                return;
            }

            try
            {
                ExportRowsToExcel(dialog.FileName);
                Status = string.Format("Отчет сохранен: {0}", dialog.FileName);
            }
            catch (Exception ex)
            {
                Status = "Ошибка выгрузки отчета";
                TaskDialog.Show("Выгрузка в Excel", ex.Message);
            }
        }

        private string GetDefaultReportFileName()
        {
            return string.Format(
                "ОтчётПоОбщПар_{0}.xlsx",
                DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
        }

        private void ExportRowsToExcel(string filePath)
        {
            List<ReportColumn> columns = GetReportColumns();
            List<ParameterUsageRow> rows = Rows.ToList();

            using (Package package = Package.Open(filePath, FileMode.Create))
            {
                Uri workbookUri = new Uri("/xl/workbook.xml", UriKind.Relative);
                Uri worksheetUri = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);

                PackagePart workbookPart = package.CreatePart(
                    workbookUri,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    CompressionOption.Maximum);

                PackagePart worksheetPart = package.CreatePart(
                    worksheetUri,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                    CompressionOption.Maximum);

                package.CreateRelationship(
                    workbookUri,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

                workbookPart.CreateRelationship(
                    new Uri("worksheets/sheet1.xml", UriKind.Relative),
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    "rId1");

                WritePartText(workbookPart, BuildWorkbookXml());
                WritePartText(worksheetPart, BuildWorksheetXml(columns, rows));
            }
        }

        private List<ReportColumn> GetReportColumns()
        {
            List<ReportColumn> columns = new List<ReportColumn>
            {
                new ReportColumn("ID", "ParameterIdText", 14),
                new ReportColumn("Имя", "Name", 34),
                new ReportColumn("GUID", "Guid", 38)
            };

            if (_doc.IsFamilyDocument)
            {
                columns.Add(new ReportColumn("Параметры с формулой", "FormulaParameters", 34));
                columns.Add(new ReportColumn("ID размера (имена видов)", "Dimensions", 38));
            }
            else
            {
                columns.Add(new ReportColumn("Категории", "Families", 42));
                columns.Add(new ReportColumn("Фильтры видов", "ViewFilters", 34));
                columns.Add(new ReportColumn("Содержится в спецификации", "Schedules", 34));
                columns.Add(new ReportColumn("Фильтры спецификаций", "ScheduleFilters", 34));
            }

            return columns;
        }

        private string BuildWorkbookXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                   "<sheets><sheet name=\"Отчет\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                   "</workbook>";
        }

        private string BuildWorksheetXml(List<ReportColumn> columns, List<ParameterUsageRow> rows)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            builder.Append("<sheetViews><sheetView workbookViewId=\"0\"/></sheetViews>");
            builder.Append("<sheetFormatPr defaultRowHeight=\"15\"/>");
            builder.Append("<cols>");

            for (int i = 0; i < columns.Count; i++)
            {
                builder.AppendFormat(
                    CultureInfo.InvariantCulture,
                    "<col min=\"{0}\" max=\"{0}\" width=\"{1}\" customWidth=\"1\"/>",
                    i + 1,
                    columns[i].Width);
            }

            builder.Append("</cols>");
            builder.Append("<sheetData>");
            AppendExcelRow(builder, 1, columns.Select(x => x.Header).ToList());

            int rowIndex = 2;
            foreach (ParameterUsageRow row in rows)
            {
                List<string> values = columns
                    .Select(x => GetColumnValue(row, x.SortMemberPath))
                    .ToList();

                AppendExcelRow(builder, rowIndex, values);
                rowIndex++;
            }

            builder.Append("</sheetData>");
            builder.Append("</worksheet>");

            return builder.ToString();
        }

        private void AppendExcelRow(StringBuilder builder, int rowIndex, List<string> values)
        {
            builder.AppendFormat(CultureInfo.InvariantCulture, "<row r=\"{0}\">", rowIndex);

            for (int i = 0; i < values.Count; i++)
            {
                AppendExcelCell(builder, rowIndex, i + 1, values[i]);
            }

            builder.Append("</row>");
        }

        private void AppendExcelCell(StringBuilder builder, int rowIndex, int columnIndex, string value)
        {
            builder.Append("<c r=\"");
            builder.Append(GetExcelColumnName(columnIndex));
            builder.Append(rowIndex.ToString(CultureInfo.InvariantCulture));
            builder.Append("\" t=\"inlineStr\"><is><t xml:space=\"preserve\">");
            builder.Append(SecurityElement.Escape(CleanXmlText(value)));
            builder.Append("</t></is></c>");
        }

        private string GetExcelColumnName(int columnIndex)
        {
            StringBuilder builder = new StringBuilder();
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                builder.Insert(0, Convert.ToChar('A' + modulo));
                columnIndex = (columnIndex - modulo) / 26;
            }

            return builder.ToString();
        }

        private string CleanXmlText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char character in value)
            {
                if (IsValidXmlCharacter(character))
                {
                    builder.Append(character);
                }
            }

            return builder.ToString();
        }

        private bool IsValidXmlCharacter(char character)
        {
            return character == 0x9 ||
                   character == 0xA ||
                   character == 0xD ||
                   (character >= 0x20 && character <= 0xD7FF) ||
                   (character >= 0xE000 && character <= 0xFFFD);
        }

        private void WritePartText(PackagePart part, string text)
        {
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(text);
            }
        }

        private void ParametersGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.C || (Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
            {
                return;
            }

            System.Windows.Forms.TextBox focusedTextBox = Keyboard.FocusedElement as System.Windows.Forms.TextBox;
            if (focusedTextBox != null && !string.IsNullOrEmpty(focusedTextBox.SelectedText))
            {
                return;
            }

            string clipboardText = BuildSelectedCellsClipboardText();
            if (!string.IsNullOrWhiteSpace(clipboardText))
            {
                Clipboard.SetText(clipboardText);
                Status = "Скопировано в буфер обмена";
                e.Handled = true;
            }
        }

        private string BuildSelectedCellsClipboardText()
        {
            List<DataGridCellInfo> selectedCells = ParametersGrid.SelectedCells
                .Where(x => x.Item as ParameterUsageRow != null &&
                            x.Column.Visibility == System.Windows.Visibility.Visible &&
                            !string.IsNullOrWhiteSpace(x.Column.SortMemberPath))
                .ToList();

            if (selectedCells.Count == 0)
            {
                return string.Empty;
            }

            List<DataGridColumn> visibleColumns = ParametersGrid.Columns
                .Where(x => x.Visibility == System.Windows.Visibility.Visible &&
                            !string.IsNullOrWhiteSpace(x.SortMemberPath) &&
                            selectedCells.Any(y => y.Column == x))
                .OrderBy(x => x.DisplayIndex)
                .ToList();

            List<ParameterUsageRow> selectedRows = Rows
                .Where(x => selectedCells.Any(y => ReferenceEquals(y.Item, x)))
                .ToList();

            StringBuilder builder = new StringBuilder();

            foreach (ParameterUsageRow row in selectedRows)
            {
                List<string> values = new List<string>();
                foreach (DataGridColumn column in visibleColumns)
                {
                    bool hasCell = selectedCells.Any(x => ReferenceEquals(x.Item, row) && x.Column == column);
                    values.Add(hasCell ? NormalizeClipboardCell(GetColumnValue(row, column.SortMemberPath)) : string.Empty);
                }

                builder.AppendLine(string.Join("\t", values));
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private string GetColumnValue(ParameterUsageRow row, string sortMemberPath)
        {
            if (sortMemberPath == "ParameterIdText")
            {
                return row.ParameterIdText;
            }

            if (sortMemberPath == "Name")
            {
                return row.Name;
            }

            if (sortMemberPath == "Guid")
            {
                return row.Guid;
            }

            if (sortMemberPath == "Families")
            {
                return row.Families;
            }

            if (sortMemberPath == "ViewFilters")
            {
                return row.ViewFilters;
            }

            if (sortMemberPath == "Schedules")
            {
                return row.Schedules;
            }

            if (sortMemberPath == "ScheduleFilters")
            {
                return row.ScheduleFilters;
            }

            if (sortMemberPath == "FormulaParameters")
            {
                return row.FormulaParameters;
            }

            if (sortMemberPath == "Dimensions")
            {
                return row.Dimensions;
            }

            return string.Empty;
        }

        private string NormalizeClipboardCell(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace("\r\n", "; ").Replace("\n", "; ").Replace("\r", "; ");
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public class ParameterUsageRow : INotifyPropertyChanged
        {
            private bool _isSelectedForDelete;
            private bool _familiesExpanded;
            private bool _viewFiltersExpanded;
            private bool _schedulesExpanded;
            private bool _scheduleFiltersExpanded;
            private bool _formulaParametersExpanded;
            private bool _dimensionsExpanded;
            private string _displayFamilies;
            private string _displayViewFilters;
            private string _displaySchedules;
            private string _displayScheduleFilters;
            private string _displayFormulaParameters;
            private string _displayDimensions;

            public event PropertyChangedEventHandler PropertyChanged;

            public string ParameterIdText { get; set; }
            public string Name { get; set; }
            public string Guid { get; set; }
            public string Families { get; set; }
            public string ViewFilters { get; set; }
            public string Schedules { get; set; }
            public string ScheduleFilters { get; set; }
            public string FormulaParameters { get; set; }
            public string Dimensions { get; set; }
            public string DisplayFamilies
            {
                get { return _displayFamilies; }
                private set { SetDisplayValue(ref _displayFamilies, value, "DisplayFamilies"); }
            }
            public string DisplayViewFilters
            {
                get { return _displayViewFilters; }
                private set { SetDisplayValue(ref _displayViewFilters, value, "DisplayViewFilters"); }
            }
            public string DisplaySchedules
            {
                get { return _displaySchedules; }
                private set { SetDisplayValue(ref _displaySchedules, value, "DisplaySchedules"); }
            }
            public string DisplayScheduleFilters
            {
                get { return _displayScheduleFilters; }
                private set { SetDisplayValue(ref _displayScheduleFilters, value, "DisplayScheduleFilters"); }
            }
            public string DisplayFormulaParameters
            {
                get { return _displayFormulaParameters; }
                private set { SetDisplayValue(ref _displayFormulaParameters, value, "DisplayFormulaParameters"); }
            }
            public string DisplayDimensions
            {
                get { return _displayDimensions; }
                private set { SetDisplayValue(ref _displayDimensions, value, "DisplayDimensions"); }
            }
            public bool IsProjectParameter { get; set; }
            public bool IsSelectedForDelete
            {
                get { return _isSelectedForDelete; }
                set
                {
                    if (_isSelectedForDelete == value)
                    {
                        return;
                    }

                    _isSelectedForDelete = value;
                    OnPropertyChanged("IsSelectedForDelete");
                }
            }

            internal Guid ParameterGuid { get; set; }
            internal ElementId ParameterId { get; set; }

            internal SortedSet<string> FamiliesSet { get; private set; }
            internal SortedSet<string> ViewFiltersSet { get; private set; }
            internal SortedSet<string> SchedulesSet { get; private set; }
            internal SortedSet<string> ScheduleFiltersSet { get; private set; }

            public ParameterUsageRow()
            {
                ParameterIdText = string.Empty;
                Name = string.Empty;
                Guid = string.Empty;
                Families = string.Empty;
                ViewFilters = string.Empty;
                Schedules = string.Empty;
                ScheduleFilters = string.Empty;
                FormulaParameters = string.Empty;
                Dimensions = string.Empty;
                _displayFamilies = string.Empty;
                _displayViewFilters = string.Empty;
                _displaySchedules = string.Empty;
                _displayScheduleFilters = string.Empty;
                _displayFormulaParameters = string.Empty;
                _displayDimensions = string.Empty;
                IsProjectParameter = false;
                _isSelectedForDelete = false;

                ParameterGuid = System.Guid.Empty;
                ParameterId = ElementId.InvalidElementId;

                FamiliesSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                ViewFiltersSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                SchedulesSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                ScheduleFiltersSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
            }

            internal void CollapseCompactCells()
            {
                _familiesExpanded = false;
                _viewFiltersExpanded = false;
                _schedulesExpanded = false;
                _scheduleFiltersExpanded = false;
                _formulaParametersExpanded = false;
                _dimensionsExpanded = false;
            }

            internal void UpdateCompactDisplay(bool isCompactMode)
            {
                DisplayFamilies = GetCompactDisplayValue(Families, isCompactMode, _familiesExpanded);
                DisplayViewFilters = GetCompactDisplayValue(ViewFilters, isCompactMode, _viewFiltersExpanded);
                DisplaySchedules = GetCompactDisplayValue(Schedules, isCompactMode, _schedulesExpanded);
                DisplayScheduleFilters = GetCompactDisplayValue(ScheduleFilters, isCompactMode, _scheduleFiltersExpanded);
                DisplayFormulaParameters = GetCompactDisplayValue(FormulaParameters, isCompactMode, _formulaParametersExpanded);
                DisplayDimensions = GetCompactDisplayValue(Dimensions, isCompactMode, _dimensionsExpanded);
            }

            internal int GetCompactItemCount(string cellKey)
            {
                return CountCompactItems(GetCompactCellValue(cellKey));
            }

            internal bool IsCompactCellCollapsed(string cellKey)
            {
                return GetCompactItemCount(cellKey) > 1 && !IsCompactCellExpanded(cellKey);
            }

            internal void ToggleCompactCell(string cellKey, bool isCompactMode)
            {
                if (!ToggleCompactCellExpanded(cellKey))
                {
                    return;
                }

                UpdateCompactCellDisplay(cellKey, isCompactMode);
            }

            private void UpdateCompactCellDisplay(string cellKey, bool isCompactMode)
            {
                if (cellKey == "Families")
                {
                    DisplayFamilies = GetCompactDisplayValue(Families, isCompactMode, _familiesExpanded);
                    return;
                }

                if (cellKey == "ViewFilters")
                {
                    DisplayViewFilters = GetCompactDisplayValue(ViewFilters, isCompactMode, _viewFiltersExpanded);
                    return;
                }

                if (cellKey == "Schedules")
                {
                    DisplaySchedules = GetCompactDisplayValue(Schedules, isCompactMode, _schedulesExpanded);
                    return;
                }

                if (cellKey == "ScheduleFilters")
                {
                    DisplayScheduleFilters = GetCompactDisplayValue(ScheduleFilters, isCompactMode, _scheduleFiltersExpanded);
                    return;
                }

                if (cellKey == "FormulaParameters")
                {
                    DisplayFormulaParameters = GetCompactDisplayValue(FormulaParameters, isCompactMode, _formulaParametersExpanded);
                    return;
                }

                if (cellKey == "Dimensions")
                {
                    DisplayDimensions = GetCompactDisplayValue(Dimensions, isCompactMode, _dimensionsExpanded);
                }
            }

            private string GetCompactCellValue(string cellKey)
            {
                if (cellKey == "Families")
                {
                    return Families;
                }

                if (cellKey == "ViewFilters")
                {
                    return ViewFilters;
                }

                if (cellKey == "Schedules")
                {
                    return Schedules;
                }

                if (cellKey == "ScheduleFilters")
                {
                    return ScheduleFilters;
                }

                if (cellKey == "FormulaParameters")
                {
                    return FormulaParameters;
                }

                if (cellKey == "Dimensions")
                {
                    return Dimensions;
                }

                return string.Empty;
            }

            private bool IsCompactCellExpanded(string cellKey)
            {
                if (cellKey == "Families")
                {
                    return _familiesExpanded;
                }

                if (cellKey == "ViewFilters")
                {
                    return _viewFiltersExpanded;
                }

                if (cellKey == "Schedules")
                {
                    return _schedulesExpanded;
                }

                if (cellKey == "ScheduleFilters")
                {
                    return _scheduleFiltersExpanded;
                }

                if (cellKey == "FormulaParameters")
                {
                    return _formulaParametersExpanded;
                }

                if (cellKey == "Dimensions")
                {
                    return _dimensionsExpanded;
                }

                return false;
            }

            private bool ToggleCompactCellExpanded(string cellKey)
            {
                if (cellKey == "Families")
                {
                    _familiesExpanded = !_familiesExpanded;
                    return true;
                }

                if (cellKey == "ViewFilters")
                {
                    _viewFiltersExpanded = !_viewFiltersExpanded;
                    return true;
                }

                if (cellKey == "Schedules")
                {
                    _schedulesExpanded = !_schedulesExpanded;
                    return true;
                }

                if (cellKey == "ScheduleFilters")
                {
                    _scheduleFiltersExpanded = !_scheduleFiltersExpanded;
                    return true;
                }

                if (cellKey == "FormulaParameters")
                {
                    _formulaParametersExpanded = !_formulaParametersExpanded;
                    return true;
                }

                if (cellKey == "Dimensions")
                {
                    _dimensionsExpanded = !_dimensionsExpanded;
                    return true;
                }

                return false;
            }

            private static string GetCompactDisplayValue(string value, bool isCompactMode, bool isExpanded)
            {
                int itemCount = CountCompactItems(value);
                if (!isCompactMode || isExpanded || itemCount <= 1)
                {
                    return value ?? string.Empty;
                }

                return string.Format("▶ Кол-во элементов ({0})", itemCount);
            }

            private static int CountCompactItems(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return 0;
                }

                return value
                    .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.RemoveEmptyEntries)
                    .Count(x => !string.IsNullOrWhiteSpace(x));
            }

            private void SetDisplayValue(ref string field, string value, string propertyName)
            {
                value = value ?? string.Empty;
                if (field == value)
                {
                    return;
                }

                field = value;
                OnPropertyChanged(propertyName);
            }

            private void OnPropertyChanged(string propertyName)
            {
                PropertyChangedEventHandler handler = PropertyChanged;
                if (handler != null)
                {
                    handler(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }

        private class ReportColumn
        {
            internal string Header { get; private set; }
            internal string SortMemberPath { get; private set; }
            internal double Width { get; private set; }

            internal ReportColumn(string header, string sortMemberPath, double width)
            {
                Header = header;
                SortMemberPath = sortMemberPath;
                Width = width;
            }
        }
    }
}