using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class CollectionParametersWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;

        private string _windowTitle;
        private string _description;
        private string _status;

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

        public CollectionParametersWindow(ExternalCommandData commandData)
        {
            _doc = commandData.Application.ActiveUIDocument.Document;
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

        private void LoadProjectParametersUsage()
        {
            List<ParameterUsageRow> parameterRows = GetBoundSharedProjectParameters();
            Dictionary<int, ParameterUsageRow> rowsByParameterId = parameterRows
                .Where(x => IsValidElementId(x.ParameterId))
                .GroupBy(x => x.ParameterId.IntegerValue)
                .ToDictionary(x => x.Key, x => x.First());

            FillViewFiltersUsage(rowsByParameterId);
            FillScheduleFiltersUsage(rowsByParameterId);

            foreach (ParameterUsageRow row in parameterRows.OrderBy(x => x.Name))
            {
                row.Families = JoinLines(row.FamiliesSet);
                row.ViewFilters = JoinLines(row.ViewFiltersSet);
                row.ScheduleFilters = JoinLines(row.ScheduleFiltersSet);
                Rows.Add(row);
            }
        }

        private List<ParameterUsageRow> GetBoundSharedProjectParameters()
        {
            List<ParameterUsageRow> result = new List<ParameterUsageRow>();
            HashSet<string> usedGuids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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

                string guidText = guid.ToString();
                if (!usedGuids.Add(guidText))
                {
                    continue;
                }

                ParameterUsageRow row = new ParameterUsageRow
                {
                    ParameterIdText = GetElementIdText(parameterId),
                    Name = definition.Name,
                    Guid = guidText,
                    ParameterGuid = guid,
                    ParameterId = parameterId
                };

                AddBindingCategoriesUsage(row, binding);
                result.Add(row);
            }

            return result;
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
            return string.Format("{0} - {1}", category.Id.IntegerValue, category.Name);
        }

        private ElementId GetSharedParameterElementId(Guid guid, string parameterName)
        {
            SharedParameterElement sharedParameter = new FilteredElementCollector(_doc)
                .OfClass(typeof(SharedParameterElement))
                .Cast<SharedParameterElement>()
                .FirstOrDefault(x => x.GuidValue == guid ||
                                     (x.GetDefinition() != null && x.GetDefinition().Name == parameterName));

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

            SharedParameterElement byName = new FilteredElementCollector(_doc)
                .OfClass(typeof(SharedParameterElement))
                .Cast<SharedParameterElement>()
                .FirstOrDefault(x => x.GetDefinition() != null && x.GetDefinition().Name == definition.Name);

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

            return string.Format("{0} - {1}", element.Id.IntegerValue, element.Name);
        }

        private string GetElementIdText(ElementId elementId)
        {
            return IsValidElementId(elementId) ? elementId.IntegerValue.ToString() : string.Empty;
        }

        private void FillViewFiltersUsage(Dictionary<int, ParameterUsageRow> rowsByParameterId)
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
                HashSet<int> parameterIds = GetParameterIdsFromViewFilter(filter);
                foreach (int parameterId in parameterIds)
                {
                    ParameterUsageRow row;
                    if (rowsByParameterId.TryGetValue(parameterId, out row))
                    {
                        row.ViewFiltersSet.Add(GetNamedElementReferenceText(filter));
                    }
                }
            }
        }

        private HashSet<int> GetParameterIdsFromViewFilter(ParameterFilterElement filter)
        {
            HashSet<int> result = new HashSet<int>();

            AddRuleParameterIds(filter, result);

            object elementFilter = InvokeParameterlessMethod(filter, "GetElementFilter");
            AddElementFilterParameterIds(elementFilter, result);

            return result;
        }

        private void AddElementFilterParameterIds(object elementFilter, HashSet<int> result)
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

        private void AddNestedElementFilters(object filtersObject, HashSet<int> result)
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

        private void AddRuleParameterIds(object source, HashSet<int> result)
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
                    result.Add(parameterId.IntegerValue);
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

        private void FillScheduleFiltersUsage(Dictionary<int, ParameterUsageRow> rowsByParameterId)
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
                        if (rowsByParameterId.TryGetValue(parameterId.IntegerValue, out row))
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

        private void LoadFamilyParametersUsage()
        {
            FamilyManager familyManager = _doc.FamilyManager;
            List<FamilyParameter> allParameters = familyManager.Parameters
                .Cast<FamilyParameter>()
                .ToList();

            List<FamilyParameter> sharedParameters = allParameters
                .Where(x => x.IsShared)
                .OrderBy(x => x.Definition.Name)
                .ToList();

            Dictionary<string, SortedSet<string>> formulasByGuid = GetFamilyFormulaUsage(allParameters, sharedParameters);
            Dictionary<string, SortedSet<string>> dimensionsByGuid = GetFamilyDimensionUsage(sharedParameters);

            foreach (FamilyParameter parameter in sharedParameters)
            {
                string guidText = parameter.GUID.ToString();

                ParameterUsageRow row = new ParameterUsageRow
                {
                    ParameterIdText = GetElementIdText(parameter.Id),
                    Name = parameter.Definition.Name,
                    Guid = guidText,
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
                string parameterName = sourceParameter.Definition.Name;
                string guidText = sourceParameter.GUID.ToString();

                foreach (FamilyParameter formulaParameter in allParameters)
                {
                    if (formulaParameter == sourceParameter || string.IsNullOrWhiteSpace(formulaParameter.Formula))
                    {
                        continue;
                    }

                    if (FormulaContainsParameter(formulaParameter.Formula, parameterName))
                    {
                        result[guidText].Add(formulaParameter.Definition.Name);
                    }
                }
            }

            return result;
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
            Dictionary<int, string> guidByParameterId = sharedParameters
                .GroupBy(x => x.Id.IntegerValue)
                .ToDictionary(x => x.Key, x => x.First().GUID.ToString());

            List<Dimension> dimensions = new FilteredElementCollector(_doc)
                .OfClass(typeof(Dimension))
                .Cast<Dimension>()
                .ToList();

            Dictionary<int, SortedSet<string>> viewNamesByDimensionId = GetDimensionViewNames(dimensions);

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
                if (!guidByParameterId.TryGetValue(label.Id.IntegerValue, out guidText))
                {
                    continue;
                }

                result[guidText].Add(GetDimensionReferenceText(dimension, viewNamesByDimensionId));
            }

            return result;
        }

        private Dictionary<int, SortedSet<string>> GetDimensionViewNames(List<Dimension> dimensions)
        {
            Dictionary<int, SortedSet<string>> result = dimensions
                .GroupBy(x => x.Id.IntegerValue)
                .ToDictionary(
                    x => x.Key,
                    x => new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase));

            foreach (Dimension dimension in dimensions)
            {
                View ownerView = GetOwnerView(dimension);
                if (ownerView != null)
                {
                    result[dimension.Id.IntegerValue].Add(ownerView.Name);
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
                        if (result.TryGetValue(dimension.Id.IntegerValue, out viewNames))
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

        private string GetDimensionReferenceText(Dimension dimension, Dictionary<int, SortedSet<string>> viewNamesByDimensionId)
        {
            SortedSet<string> viewNames;
            if (viewNamesByDimensionId.TryGetValue(dimension.Id.IntegerValue, out viewNames) && viewNames.Count > 0)
            {
                return string.Format(
                    "{0} ({1})",
                    dimension.Id.IntegerValue,
                    string.Join("; ", viewNames));
            }

            return string.Format("{0} (вид не найден)", dimension.Id.IntegerValue);
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
                   elementId.IntegerValue > 0;
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
                .Where(x => x.Item as ParameterUsageRow != null && x.Column.Visibility == System.Windows.Visibility.Visible)
                .ToList();

            if (selectedCells.Count == 0)
            {
                return string.Empty;
            }

            List<DataGridColumn> visibleColumns = ParametersGrid.Columns
                .Where(x => x.Visibility == System.Windows.Visibility.Visible && selectedCells.Any(y => y.Column == x))
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

        public class ParameterUsageRow
        {
            public string ParameterIdText { get; set; }
            public string Name { get; set; }
            public string Guid { get; set; }
            public string Families { get; set; }
            public string ViewFilters { get; set; }
            public string ScheduleFilters { get; set; }
            public string FormulaParameters { get; set; }
            public string Dimensions { get; set; }

            internal Guid ParameterGuid { get; set; }
            internal ElementId ParameterId { get; set; }

            internal SortedSet<string> FamiliesSet { get; private set; }
            internal SortedSet<string> ViewFiltersSet { get; private set; }
            internal SortedSet<string> ScheduleFiltersSet { get; private set; }

            public ParameterUsageRow()
            {
                ParameterIdText = string.Empty;
                Name = string.Empty;
                Guid = string.Empty;
                Families = string.Empty;
                ViewFilters = string.Empty;
                ScheduleFilters = string.Empty;
                FormulaParameters = string.Empty;
                Dimensions = string.Empty;

                ParameterGuid = System.Guid.Empty;
                ParameterId = ElementId.InvalidElementId;

                FamiliesSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                ViewFiltersSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
                ScheduleFiltersSet = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
            }
        }
    }
}
