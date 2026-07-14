using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Tools.Common;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TextBox = System.Windows.Controls.TextBox;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public partial class CollectionFamilyParametersWindow : Window, INotifyPropertyChanged
    {
        private readonly Document _doc;
        private string _windowTitle;
        private string _description;
        private string _status;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<FamilyParameterUsageRow> Rows { get; private set; }

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

        public CollectionFamilyParametersWindow(ExternalCommandData commandData)
        {
            _doc = commandData.Application.ActiveUIDocument.Document;
            Rows = new ObservableCollection<FamilyParameterUsageRow>();

            InitializeComponent();
            DataContext = this;

            WindowTitle = "KPLN. Анализ параметров семейства";
            Description = "Параметры семейства";

            LoadData();
        }

        private void LoadData()
        {
            string stage = "инициализация";

            try
            {
                Rows.Clear();
                Status = "Идет сбор данных...";

                stage = "проверка документа";
                if (!_doc.IsFamilyDocument)
                {
                    throw new InvalidOperationException(
                        "Команда доступна только в редакторе семейств.");
                }

                stage = "получение параметров FamilyManager";
                List<FamilyParameter> parameters = _doc.FamilyManager.Parameters
                    .Cast<FamilyParameter>()
                    .ToList();

                stage = "создание строк таблицы";
                List<FamilyParameterUsageRow> parameterRows =
                    CreateParameterRows(parameters);

                Dictionary<long, FamilyParameterUsageRow> rowsByParameterId =
                    parameterRows
                        .Where(x => IsUsableElementId(x.ParameterId))
                        .GroupBy(x => IDHelper.ElIdValue(x.ParameterId))
                        .ToDictionary(x => x.Key, x => x.First());

                stage = "анализ параметров выбора типоразмера вложенных семейств";
                FillNestedFamilyTypeParameterValues(
                    parameters,
                    rowsByParameterId);

                stage = "анализ параметров вложенных семейств";
                FillNestedFamilyParameterAssociations(rowsByParameterId);

                stage = "анализ привязок системных элементов";
                FillAssociatedElementParameters(parameters, rowsByParameterId);

                stage = "анализ формул текущего семейства";
                FillCurrentFamilyFormulaUsage(parameters, rowsByParameterId);

                stage = "формирование таблицы";
                foreach (FamilyParameterUsageRow row in
                         parameterRows.OrderBy(x => x.Name))
                {
                    row.FinalizeValues();
                    Rows.Add(row);
                }

                int formulaReferenceCount =
                    Rows.Sum(x => x.CurrentFamilyReferenceCount);
                int nestedBindingCount =
                    Rows.Sum(x => x.NestedFamilyBindingCount);

                Status = string.Format(
                    CultureInfo.CurrentCulture,
                    "Найдено параметров: {0}; привязок: {1}; ссылок в формулах: {2}",
                    Rows.Count,
                    nestedBindingCount,
                    formulaReferenceCount);
            }
            catch (Exception ex)
            {
                Status = "Ошибка сбора данных на этапе: " + stage;
                TaskDialog.Show(
                    "Анализ параметров семейства",
                    "Этап: " + stage + Environment.NewLine +
                    Environment.NewLine + ex);
            }
        }

        private List<FamilyParameterUsageRow> CreateParameterRows(
            List<FamilyParameter> parameters)
        {
            List<FamilyParameterUsageRow> result =
                new List<FamilyParameterUsageRow>();

            foreach (FamilyParameter parameter in parameters)
            {
                if (parameter == null)
                {
                    continue;
                }

                FamilyParameterUsageRow row = new FamilyParameterUsageRow
                {
                    ParameterIdText = GetElementIdText(parameter.Id),
                    Name = GetFamilyParameterName(parameter),
                    Guid = GetFamilyParameterGuid(parameter),
                    Formula = GetFamilyParameterFormula(parameter),
                    ParameterId = parameter.Id
                };

                result.Add(row);
            }

            return result;
        }

        private void FillCurrentFamilyFormulaUsage(
            List<FamilyParameter> parameters,
            Dictionary<long, FamilyParameterUsageRow> rowsByParameterId)
        {
            List<FormulaSource> formulaSources = parameters
                .Where(x => x != null && IsUsableElementId(x.Id))
                .Select(x => new FormulaSource(
                    x.Id,
                    GetFamilyParameterName(x),
                    GetFamilyParameterFormula(x)))
                .Where(x => !string.IsNullOrWhiteSpace(x.Formula))
                .ToList();

            foreach (FamilyParameter referencedParameter in parameters)
            {
                if (referencedParameter == null ||
                    !IsUsableElementId(referencedParameter.Id))
                {
                    continue;
                }

                FamilyParameterUsageRow row;
                if (!rowsByParameterId.TryGetValue(
                        IDHelper.ElIdValue(referencedParameter.Id),
                        out row))
                {
                    continue;
                }

                string referencedName =
                    GetFamilyParameterName(referencedParameter);

                if (string.IsNullOrWhiteSpace(referencedName))
                {
                    continue;
                }

                foreach (FormulaSource formulaSource in formulaSources)
                {
                    if (IDHelper.ElIdValue(formulaSource.ParameterId) ==
                        IDHelper.ElIdValue(referencedParameter.Id))
                    {
                        continue;
                    }

                    if (!FormulaContainsParameter(
                            formulaSource.Formula,
                            referencedName))
                    {
                        continue;
                    }

                    string target = string.Format(
                        CultureInfo.CurrentCulture,
                        "{0} [{1}]",
                        formulaSource.ParameterName,
                        GetElementIdText(formulaSource.ParameterId));

                    row.AddCurrentFamilyFormulaReference(
                        new ParameterReferenceItem(
                            formulaSource.Formula,
                            target,
                            Brushes.RoyalBlue));
                }
            }
        }

        private void FillNestedFamilyTypeParameterValues(
            List<FamilyParameter> familyParameters,
            Dictionary<long, FamilyParameterUsageRow> rowsByParameterId)
        {
            if (familyParameters == null ||
                rowsByParameterId == null ||
                rowsByParameterId.Count == 0)
            {
                return;
            }

            FamilyManager familyManager = _doc.FamilyManager;
            List<FamilyType> familyTypes = GetFamilyTypes(familyManager);
            if (familyTypes.Count == 0)
            {
                return;
            }

            foreach (FamilyParameter familyParameter in familyParameters)
            {
                if (familyParameter == null ||
                    familyParameter.StorageType != StorageType.ElementId ||
                    !IsUsableElementId(familyParameter.Id))
                {
                    continue;
                }

                FamilyParameterUsageRow row;
                if (!rowsByParameterId.TryGetValue(
                        IDHelper.ElIdValue(familyParameter.Id),
                        out row))
                {
                    continue;
                }

                foreach (FamilyType familyType in familyTypes)
                {
                    ElementId selectedTypeId =
                        GetFamilyTypeParameterElementId(
                            familyType,
                            familyParameter);

                    if (!IsUsableElementId(selectedTypeId))
                    {
                        continue;
                    }

                    string nestedFamilyName;
                    string nestedTypeName;
                    if (!TryGetNestedFamilyTypeNames(
                            selectedTypeId,
                            out nestedFamilyName,
                            out nestedTypeName))
                    {
                        continue;
                    }

                    string sourceText = string.Format(
                        CultureInfo.CurrentCulture,
                        "Для типоразмера вложенного семейства \"{0}\"",
                        GetFamilyTypeName(familyType));

                    string selectedTypeName =
                        BuildNestedFamilyTypeDisplayName(
                            nestedFamilyName,
                            nestedTypeName);

                    string target = string.Format(
                        CultureInfo.CurrentCulture,
                        "{0} [{1}]",
                        selectedTypeName,
                        GetElementIdText(selectedTypeId));

                    row.AddNestedFamilyBinding(
                        new ParameterReferenceItem(
                            sourceText,
                            target,
                            Brushes.MediumPurple));
                }
            }
        }

        private List<FamilyType> GetFamilyTypes(
            FamilyManager familyManager)
        {
            List<FamilyType> result = new List<FamilyType>();
            if (familyManager == null)
            {
                return result;
            }

            try
            {
                foreach (FamilyType familyType in familyManager.Types)
                {
                    if (familyType != null)
                    {
                        result.Add(familyType);
                    }
                }
            }
            catch
            {
            }

            if (result.Count == 0)
            {
                try
                {
                    if (familyManager.CurrentType != null)
                    {
                        result.Add(familyManager.CurrentType);
                    }
                }
                catch
                {
                }
            }

            return result
                .GroupBy(
                    x => GetFamilyTypeName(x),
                    StringComparer.CurrentCultureIgnoreCase)
                .Select(x => x.First())
                .OrderBy(
                    x => GetFamilyTypeName(x),
                    StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private ElementId GetFamilyTypeParameterElementId(
            FamilyType familyType,
            FamilyParameter familyParameter)
        {
            if (familyType == null || familyParameter == null)
            {
                return ElementId.InvalidElementId;
            }

            try
            {
                return familyType.AsElementId(familyParameter);
            }
            catch
            {
                return ElementId.InvalidElementId;
            }
        }

        private bool TryGetNestedFamilyTypeNames(
            ElementId selectedTypeId,
            out string familyName,
            out string typeName)
        {
            familyName = string.Empty;
            typeName = string.Empty;

            if (!IsUsableElementId(selectedTypeId))
            {
                return false;
            }

            Element selectedElement;
            try
            {
                selectedElement = _doc.GetElement(selectedTypeId);
            }
            catch
            {
                return false;
            }

            if (selectedElement == null)
            {
                return false;
            }

            FamilySymbol familySymbol = selectedElement as FamilySymbol;
            if (familySymbol != null)
            {
                try
                {
                    familyName = familySymbol.Family != null
                        ? familySymbol.Family.Name ?? string.Empty
                        : string.Empty;
                }
                catch
                {
                    familyName = string.Empty;
                }

                try
                {
                    typeName = familySymbol.Name ?? string.Empty;
                }
                catch
                {
                    typeName = string.Empty;
                }

                return !string.IsNullOrWhiteSpace(familyName) ||
                       !string.IsNullOrWhiteSpace(typeName);
            }

            // Для неразделяемых вложенных семейств Revit возвращает
            // NestedFamilyTypeReference. Обращаемся к его свойствам
            // через reflection, чтобы код сохранял совместимость между
            // несколькими версиями Revit API.
            if (!string.Equals(
                    selectedElement.GetType().Name,
                    "NestedFamilyTypeReference",
                    StringComparison.Ordinal))
            {
                return false;
            }

            familyName = GetStringProperty(
                selectedElement,
                "FamilyName");
            typeName = GetStringProperty(
                selectedElement,
                "TypeName");

            if (string.IsNullOrWhiteSpace(typeName))
            {
                try
                {
                    typeName = selectedElement.Name ?? string.Empty;
                }
                catch
                {
                    typeName = string.Empty;
                }
            }

            return !string.IsNullOrWhiteSpace(familyName) ||
                   !string.IsNullOrWhiteSpace(typeName);
        }

        private string GetStringProperty(
            object source,
            string propertyName)
        {
            if (source == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            try
            {
                System.Reflection.PropertyInfo property =
                    source.GetType().GetProperty(propertyName);

                if (property == null)
                {
                    return string.Empty;
                }

                object value = property.GetValue(source, null);
                return value != null
                    ? Convert.ToString(
                        value,
                        CultureInfo.CurrentCulture) ?? string.Empty
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetFamilyTypeName(FamilyType familyType)
        {
            if (familyType == null)
            {
                return "(без имени)";
            }

            try
            {
                return !string.IsNullOrWhiteSpace(familyType.Name)
                    ? familyType.Name
                    : "(без имени)";
            }
            catch
            {
                return "(без имени)";
            }
        }

        private string BuildNestedFamilyTypeDisplayName(
            string familyName,
            string typeName)
        {
            familyName = familyName ?? string.Empty;
            typeName = typeName ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(familyName) &&
                !string.IsNullOrWhiteSpace(typeName))
            {
                return familyName + ": " + typeName;
            }

            if (!string.IsNullOrWhiteSpace(typeName))
            {
                return typeName;
            }

            if (!string.IsNullOrWhiteSpace(familyName))
            {
                return familyName;
            }

            return "Вложенный типоразмер";
        }

        private void FillNestedFamilyParameterAssociations(
            Dictionary<long, FamilyParameterUsageRow> rowsByParameterId)
        {
            FamilyManager familyManager = _doc.FamilyManager;

            List<FamilyInstance> familyInstances =
                new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilyInstance))
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();

            foreach (FamilyInstance familyInstance in familyInstances)
            {
                if (familyInstance == null ||
                    !IsUsableElementId(familyInstance.Id))
                {
                    continue;
                }

                // Параметр выбора типоразмера проверяем явно: в некоторых
                // версиях Revit он может не попасть в обычный ParameterSet.
                AddNestedFamilyParameterAssociation(
                    familyManager,
                    familyInstance,
                    GetBuiltInParameter(
                        familyInstance,
                        BuiltInParameter.ELEM_TYPE_PARAM),
                    false,
                    rowsByParameterId);

                AddNestedFamilyParameterAssociation(
                    familyManager,
                    familyInstance,
                    GetBuiltInParameter(
                        familyInstance,
                        BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM),
                    false,
                    rowsByParameterId);

                foreach (Parameter instanceParameter in
                         GetElementParameters(familyInstance))
                {
                    AddNestedFamilyParameterAssociation(
                        familyManager,
                        familyInstance,
                        instanceParameter,
                        false,
                        rowsByParameterId);
                }

                FamilySymbol symbol = GetFamilySymbol(familyInstance);
                if (symbol == null)
                {
                    continue;
                }

                foreach (Parameter typeParameter in
                         GetElementParameters(symbol))
                {
                    AddNestedFamilyParameterAssociation(
                        familyManager,
                        familyInstance,
                        typeParameter,
                        true,
                        rowsByParameterId);
                }
            }
        }

        private void AddNestedFamilyParameterAssociation(
            FamilyManager familyManager,
            FamilyInstance familyInstance,
            Parameter nestedParameter,
            bool isTypeParameter,
            Dictionary<long, FamilyParameterUsageRow> rowsByParameterId)
        {
            if (familyManager == null ||
                familyInstance == null ||
                nestedParameter == null)
            {
                return;
            }

            FamilyParameter controllingParameter;
            try
            {
                controllingParameter =
                    familyManager.GetAssociatedFamilyParameter(
                        nestedParameter);
            }
            catch
            {
                return;
            }

            if (controllingParameter == null ||
                !IsUsableElementId(controllingParameter.Id))
            {
                return;
            }

            FamilyParameterUsageRow row;
            if (!rowsByParameterId.TryGetValue(
                    IDHelper.ElIdValue(controllingParameter.Id),
                    out row))
            {
                return;
            }

            bool isTypeSelector =
                !isTypeParameter &&
                IsNestedFamilyTypeSelector(
                    familyInstance,
                    nestedParameter);

            string bindingName;
            if (isTypeSelector)
            {
                bindingName =
                    "Выбор типоразмера вложенного семейства";
            }
            else
            {
                bindingName = GetElementParameterName(nestedParameter);
                if (string.IsNullOrWhiteSpace(bindingName))
                {
                    bindingName = "Параметр без имени";
                }

                bindingName += isTypeParameter
                    ? " (параметр типа)"
                    : " (параметр экземпляра)";
            }

            string target = string.Format(
                CultureInfo.CurrentCulture,
                "{0} [{1}]",
                GetElementDisplayName(familyInstance),
                GetElementIdText(familyInstance.Id));

            row.AddNestedFamilyBinding(
                new ParameterReferenceItem(
                    bindingName,
                    target,
                    Brushes.MediumPurple));
        }

        private List<Parameter> GetElementParameters(Element element)
        {
            List<Parameter> result = new List<Parameter>();
            if (element == null)
            {
                return result;
            }

            try
            {
                ParameterSet parameters = element.Parameters;
                if (parameters == null)
                {
                    return result;
                }

                foreach (Parameter parameter in parameters)
                {
                    if (parameter != null)
                    {
                        result.Add(parameter);
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        private Parameter GetBuiltInParameter(
            Element element,
            BuiltInParameter builtInParameter)
        {
            if (element == null)
            {
                return null;
            }

            try
            {
                return element.get_Parameter(builtInParameter);
            }
            catch
            {
                return null;
            }
        }

        private FamilySymbol GetFamilySymbol(
            FamilyInstance familyInstance)
        {
            if (familyInstance == null)
            {
                return null;
            }

            try
            {
                return familyInstance.Symbol;
            }
            catch
            {
                return null;
            }
        }

        private void FillAssociatedElementParameters(
            List<FamilyParameter> familyParameters,
            Dictionary<long, FamilyParameterUsageRow> rowsByParameterId)
        {
            foreach (FamilyParameter familyParameter in familyParameters)
            {
                if (familyParameter == null ||
                    !IsUsableElementId(familyParameter.Id))
                {
                    continue;
                }

                FamilyParameterUsageRow row;
                if (!rowsByParameterId.TryGetValue(
                        IDHelper.ElIdValue(familyParameter.Id),
                        out row))
                {
                    continue;
                }

                ParameterSet associatedParameters;
                try
                {
                    associatedParameters =
                        familyParameter.AssociatedParameters;
                }
                catch
                {
                    continue;
                }

                if (associatedParameters == null)
                {
                    continue;
                }

                foreach (Parameter associatedParameter in
                         associatedParameters)
                {
                    AddAssociatedSystemElementParameter(
                        row,
                        associatedParameter);
                }
            }
        }

        private void AddAssociatedSystemElementParameter(
            FamilyParameterUsageRow row,
            Parameter associatedParameter)
        {
            if (row == null || associatedParameter == null)
            {
                return;
            }

            Element ownerElement;
            try
            {
                ownerElement = associatedParameter.Element;
            }
            catch
            {
                return;
            }

            if (ownerElement == null ||
                !IsUsableElementId(ownerElement.Id))
            {
                return;
            }

            // Вложенные FamilyInstance и FamilySymbol уже обработаны
            // прямым обходом параметров, где известен конкретный экземпляр.
            if (ownerElement is FamilyInstance ||
                ownerElement is FamilySymbol)
            {
                return;
            }

            string bindingName =
                GetElementParameterName(associatedParameter);

            if (string.IsNullOrWhiteSpace(bindingName))
            {
                bindingName = "Параметр без имени";
            }

            string target = string.Format(
                CultureInfo.CurrentCulture,
                "{0} [{1}]",
                GetElementDisplayName(ownerElement),
                GetElementIdText(ownerElement.Id));

            row.AddNestedFamilyBinding(
                new ParameterReferenceItem(
                    bindingName,
                    target,
                    Brushes.MediumPurple));
        }

        private bool IsNestedFamilyTypeSelector(
            FamilyInstance familyInstance,
            Parameter associatedParameter)
        {
            if (familyInstance == null ||
                associatedParameter == null)
            {
                return false;
            }

            Parameter typeSelectorParameter = GetBuiltInParameter(
                familyInstance,
                BuiltInParameter.ELEM_TYPE_PARAM);

            Parameter familyAndTypeParameter = GetBuiltInParameter(
                familyInstance,
                BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM);

            return HaveSameParameterId(
                       associatedParameter,
                       typeSelectorParameter) ||
                   HaveSameParameterId(
                       associatedParameter,
                       familyAndTypeParameter);
        }

        private bool HaveSameParameterId(
            Parameter first,
            Parameter second)
        {
            if (first == null || second == null ||
                first.Id == null || second.Id == null)
            {
                return false;
            }

            return IDHelper.ElIdValue(first.Id) ==
                   IDHelper.ElIdValue(second.Id);
        }

        private string GetElementDisplayName(Element element)
        {
            FamilyInstance familyInstance = element as FamilyInstance;
            if (familyInstance != null)
            {
                FamilySymbol symbol = null;
                try
                {
                    symbol = familyInstance.Symbol;
                }
                catch
                {
                }

                string familyName = string.Empty;
                string typeName = string.Empty;

                if (symbol != null)
                {
                    try
                    {
                        familyName = symbol.Family != null
                            ? symbol.Family.Name
                            : string.Empty;
                    }
                    catch
                    {
                    }

                    try
                    {
                        typeName = symbol.Name ?? string.Empty;
                    }
                    catch
                    {
                    }
                }

                if (!string.IsNullOrWhiteSpace(familyName) &&
                    !string.IsNullOrWhiteSpace(typeName) &&
                    !string.Equals(
                        familyName,
                        typeName,
                        StringComparison.CurrentCultureIgnoreCase))
                {
                    return familyName + ": " + typeName;
                }

                if (!string.IsNullOrWhiteSpace(familyName))
                {
                    return familyName;
                }

                if (!string.IsNullOrWhiteSpace(typeName))
                {
                    return typeName;
                }
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(element.Name))
                {
                    return element.Name;
                }
            }
            catch
            {
            }

            try
            {
                if (element.Category != null &&
                    !string.IsNullOrWhiteSpace(element.Category.Name))
                {
                    return element.Category.Name;
                }
            }
            catch
            {
            }

            return element.GetType().Name;
        }

        private string GetElementParameterName(Parameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                return parameter.Definition != null
                    ? parameter.Definition.Name ?? string.Empty
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetFamilyParameterName(FamilyParameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                return parameter.Definition != null
                    ? parameter.Definition.Name ?? string.Empty
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private string GetFamilyParameterGuid(FamilyParameter parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                if (!parameter.IsShared ||
                    parameter.GUID == System.Guid.Empty)
                {
                    return string.Empty;
                }

                return parameter.GUID.ToString();
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

        private bool FormulaContainsParameter(
            string formula,
            string parameterName)
        {
            if (string.IsNullOrWhiteSpace(formula) ||
                string.IsNullOrWhiteSpace(parameterName))
            {
                return false;
            }

            string pattern = string.Format(
                CultureInfo.InvariantCulture,
                @"(?<![\p{{L}}\p{{N}}_]){0}(?![\p{{L}}\p{{N}}_])",
                Regex.Escape(parameterName));

            return Regex.IsMatch(
                formula,
                pattern,
                RegexOptions.IgnoreCase |
                RegexOptions.CultureInvariant);
        }

        private string GetElementIdText(ElementId elementId)
        {
            return IsUsableElementId(elementId)
                ? IDHelper.ElIdValue(elementId).ToString(
                    CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private bool IsUsableElementId(ElementId elementId)
        {
            return elementId != null &&
                   elementId != ElementId.InvalidElementId &&
                   IDHelper.ElIdValue(elementId) > 0;
        }

        private void ExportExcelButton_Click(
            object sender,
            RoutedEventArgs e)
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
                Status = "Отчет сохранен: " + dialog.FileName;
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
                CultureInfo.InvariantCulture,
                "ОтчётПоПараметрамСемейства_{0}.xlsx",
                DateTime.Now.ToString(
                    "yyyyMMdd_HHmmss",
                    CultureInfo.InvariantCulture));
        }

        private void ExportRowsToExcel(string filePath)
        {
            List<ReportColumn> columns = new List<ReportColumn>
            {
                new ReportColumn("ID", "ParameterIdText", 14),
                new ReportColumn("Имя", "Name", 34),
                new ReportColumn("GUID", "Guid", 38),
                new ReportColumn(
                    "Привязка параметра к вложенным семействам",
                    "NestedFamilyBindingsPlainText",
                    58),
                new ReportColumn(
                    "Формула связанная с основным семейством",
                    "CurrentFamilyFormulaUsagePlainText",
                    58),
                new ReportColumn("Формула", "Formula", 48)
            };

            List<FamilyParameterUsageRow> rows = Rows.ToList();

            using (Package package = Package.Open(filePath, FileMode.Create))
            {
                Uri workbookUri =
                    new Uri("/xl/workbook.xml", UriKind.Relative);
                Uri worksheetUri = new Uri(
                    "/xl/worksheets/sheet1.xml",
                    UriKind.Relative);

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
                WritePartText(
                    worksheetPart,
                    BuildWorksheetXml(columns, rows));
            }
        }

        private string BuildWorkbookXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                   "<sheets><sheet name=\"Отчет\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                   "</workbook>";
        }

        private string BuildWorksheetXml(
            List<ReportColumn> columns,
            List<FamilyParameterUsageRow> rows)
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
            AppendExcelRow(
                builder,
                1,
                columns.Select(x => x.Header).ToList());

            int rowIndex = 2;
            foreach (FamilyParameterUsageRow row in rows)
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

        private void AppendExcelRow(
            StringBuilder builder,
            int rowIndex,
            List<string> values)
        {
            builder.AppendFormat(
                CultureInfo.InvariantCulture,
                "<row r=\"{0}\">",
                rowIndex);

            for (int i = 0; i < values.Count; i++)
            {
                AppendExcelCell(builder, rowIndex, i + 1, values[i]);
            }

            builder.Append("</row>");
        }

        private void AppendExcelCell(
            StringBuilder builder,
            int rowIndex,
            int columnIndex,
            string value)
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
            using (Stream stream = part.GetStream(
                FileMode.Create,
                FileAccess.Write))
            using (StreamWriter writer = new StreamWriter(
                stream,
                new UTF8Encoding(false)))
            {
                writer.Write(text);
            }
        }

        private void ParametersGrid_PreviewKeyDown(
            object sender,
            KeyEventArgs e)
        {
            if (e.Key != Key.C ||
                (Keyboard.Modifiers & ModifierKeys.Control) !=
                ModifierKeys.Control)
            {
                return;
            }

            TextBox focusedTextBox =
                Keyboard.FocusedElement as TextBox;

            if (focusedTextBox != null &&
                !string.IsNullOrEmpty(focusedTextBox.SelectedText))
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
            List<DataGridCellInfo> selectedCells =
                ParametersGrid.SelectedCells
                    .Where(x =>
                        x.Item as FamilyParameterUsageRow != null &&
                        x.Column.Visibility ==
                            System.Windows.Visibility.Visible &&
                        !string.IsNullOrWhiteSpace(
                            x.Column.SortMemberPath))
                    .ToList();

            if (selectedCells.Count == 0)
            {
                return string.Empty;
            }

            List<DataGridColumn> visibleColumns =
                ParametersGrid.Columns
                    .Where(x =>
                        x.Visibility ==
                            System.Windows.Visibility.Visible &&
                        !string.IsNullOrWhiteSpace(x.SortMemberPath) &&
                        selectedCells.Any(y => y.Column == x))
                    .OrderBy(x => x.DisplayIndex)
                    .ToList();

            HashSet<FamilyParameterUsageRow> selectedRowsSet =
                new HashSet<FamilyParameterUsageRow>(
                    selectedCells
                        .Select(x => x.Item)
                        .OfType<FamilyParameterUsageRow>());

            List<FamilyParameterUsageRow> selectedRows =
                ParametersGrid.Items
                    .Cast<object>()
                    .OfType<FamilyParameterUsageRow>()
                    .Where(x => selectedRowsSet.Contains(x))
                    .ToList();

            StringBuilder builder = new StringBuilder();

            foreach (FamilyParameterUsageRow row in selectedRows)
            {
                List<string> values = new List<string>();
                foreach (DataGridColumn column in visibleColumns)
                {
                    bool hasCell = selectedCells.Any(
                        x => ReferenceEquals(x.Item, row) &&
                             x.Column == column);

                    values.Add(hasCell
                        ? NormalizeClipboardCell(
                            GetColumnValue(
                                row,
                                column.SortMemberPath))
                        : string.Empty);
                }

                builder.AppendLine(string.Join("\t", values));
            }

            return builder.ToString().TrimEnd('\r', '\n');
        }

        private string GetColumnValue(
            FamilyParameterUsageRow row,
            string sortMemberPath)
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

            if (sortMemberPath == "NestedFamilyBindingsPlainText")
            {
                return row.NestedFamilyBindingsPlainText;
            }

            if (sortMemberPath == "CurrentFamilyFormulaUsagePlainText")
            {
                return row.CurrentFamilyFormulaUsagePlainText;
            }

            if (sortMemberPath == "Formula")
            {
                return row.Formula;
            }

            return string.Empty;
        }

        private string NormalizeClipboardCell(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("\r\n", "; ")
                .Replace("\n", "; ")
                .Replace("\r", "; ");
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(
                    this,
                    new PropertyChangedEventArgs(propertyName));
            }
        }

        public class FamilyParameterUsageRow
        {
            private readonly Dictionary<string, ParameterReferenceItem>
                _nestedFamilyBindings;
            private readonly Dictionary<string, ParameterReferenceItem>
                _currentFamilyFormulaUsage;

            public string ParameterIdText { get; set; }
            public string Name { get; set; }
            public string Guid { get; set; }
            public string Formula { get; set; }
            public string NestedFamilyBindingsPlainText { get; private set; }
            public string CurrentFamilyFormulaUsagePlainText { get; private set; }

            public ObservableCollection<ParameterReferenceItem>
                NestedFamilyBindings
            { get; private set; }

            public ObservableCollection<ParameterReferenceItem>
                CurrentFamilyFormulaUsage
            { get; private set; }

            public int NestedFamilyBindingCount
            {
                get { return _nestedFamilyBindings.Count; }
            }

            public int CurrentFamilyReferenceCount
            {
                get { return _currentFamilyFormulaUsage.Count; }
            }

            internal ElementId ParameterId { get; set; }

            public FamilyParameterUsageRow()
            {
                ParameterIdText = string.Empty;
                Name = string.Empty;
                Guid = string.Empty;
                Formula = string.Empty;
                NestedFamilyBindingsPlainText = string.Empty;
                CurrentFamilyFormulaUsagePlainText = string.Empty;
                ParameterId = ElementId.InvalidElementId;

                NestedFamilyBindings =
                    new ObservableCollection<ParameterReferenceItem>();
                CurrentFamilyFormulaUsage =
                    new ObservableCollection<ParameterReferenceItem>();

                _nestedFamilyBindings =
                    new Dictionary<string, ParameterReferenceItem>(
                        StringComparer.OrdinalIgnoreCase);
                _currentFamilyFormulaUsage =
                    new Dictionary<string, ParameterReferenceItem>(
                        StringComparer.OrdinalIgnoreCase);
            }

            internal void AddNestedFamilyBinding(
                ParameterReferenceItem item)
            {
                AddReferenceItem(_nestedFamilyBindings, item);
            }

            internal void AddCurrentFamilyFormulaReference(
                ParameterReferenceItem item)
            {
                AddReferenceItem(_currentFamilyFormulaUsage, item);
            }

            private void AddReferenceItem(
                Dictionary<string, ParameterReferenceItem> destination,
                ParameterReferenceItem item)
            {
                if (item == null ||
                    string.IsNullOrWhiteSpace(item.Text))
                {
                    return;
                }

                string key = item.Text + "|" + item.Target;
                if (!destination.ContainsKey(key))
                {
                    destination.Add(key, item);
                }
            }

            internal void FinalizeValues()
            {
                string nestedPlainText;
                FillReferenceCollection(
                    _nestedFamilyBindings,
                    NestedFamilyBindings,
                    out nestedPlainText);
                NestedFamilyBindingsPlainText = nestedPlainText;

                string formulaPlainText;
                FillReferenceCollection(
                    _currentFamilyFormulaUsage,
                    CurrentFamilyFormulaUsage,
                    out formulaPlainText);
                CurrentFamilyFormulaUsagePlainText = formulaPlainText;
            }

            private static void FillReferenceCollection(
                Dictionary<string, ParameterReferenceItem> source,
                ObservableCollection<ParameterReferenceItem> destination,
                out string plainText)
            {
                List<ParameterReferenceItem> orderedItems = source.Values
                    .OrderBy(
                        x => x.Text,
                        StringComparer.CurrentCultureIgnoreCase)
                    .ThenBy(
                        x => x.Target,
                        StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                destination.Clear();
                foreach (ParameterReferenceItem item in orderedItems)
                {
                    destination.Add(item);
                }

                plainText = string.Join(
                    Environment.NewLine,
                    orderedItems.Select(x => x.PlainText));
            }
        }

        public class ParameterReferenceItem
        {
            public string Text { get; private set; }
            public string Separator { get; private set; }
            public string Target { get; private set; }
            public Brush Foreground { get; private set; }

            public string PlainText
            {
                get { return Text + Separator + Target; }
            }

            internal ParameterReferenceItem(
                string text,
                string target,
                Brush foreground)
            {
                Text = text ?? string.Empty;
                Target = target ?? string.Empty;
                Separator = string.IsNullOrWhiteSpace(Target)
                    ? string.Empty
                    : " — ";
                Foreground = foreground ?? Brushes.Black;
            }
        }

        private class FormulaSource
        {
            internal ElementId ParameterId { get; private set; }
            internal string ParameterName { get; private set; }
            internal string Formula { get; private set; }

            internal FormulaSource(
                ElementId parameterId,
                string parameterName,
                string formula)
            {
                ParameterId = parameterId;
                ParameterName = parameterName ?? string.Empty;
                Formula = formula ?? string.Empty;
            }
        }

        private class ReportColumn
        {
            internal string Header { get; private set; }
            internal string SortMemberPath { get; private set; }
            internal double Width { get; private set; }

            internal ReportColumn(
                string header,
                string sortMemberPath,
                double width)
            {
                Header = header;
                SortMemberPath = sortMemberPath;
                Width = width;
            }
        }
    }
}
