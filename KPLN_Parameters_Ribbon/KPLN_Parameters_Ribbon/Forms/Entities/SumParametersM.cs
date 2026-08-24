using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Data;

namespace KPLN_Parameters_Ribbon.Forms.Entities
{
    public sealed class SumParametersM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private int _roundDigits = 2;
        private string _coefficientText = string.Empty;
        private string _searchText;
        private ICollectionView _filteredSumResults;
        private Element[] _userSelElems = new Element[0];
        private static readonly Dictionary<string, string> UnitNameAbbreviations = new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase)
        {
            { "мм", "мм" },
            { "миллиметр", "мм" },
            { "миллиметры", "мм" },
            { "миллиметров", "мм" },
            { "mm", "мм" },
            { "см", "см" },
            { "сантиметр", "см" },
            { "сантиметры", "см" },
            { "сантиметров", "см" },
            { "cm", "см" },
            { "м", "м" },
            { "метр", "м" },
            { "метры", "м" },
            { "метров", "м" },
            { "m", "м" },
            { "км", "км" },
            { "километр", "км" },
            { "километры", "км" },
            { "километров", "км" },
            { "km", "км" },
            { "мм²", "мм²" },
            { "mm²", "мм²" },
            { "mm2", "мм²" },
            { "кв. мм", "мм²" },
            { "квадратные миллиметры", "мм²" },
            { "square millimeters", "мм²" },
            { "см²", "см²" },
            { "cm²", "см²" },
            { "cm2", "см²" },
            { "кв. см", "см²" },
            { "квадратные сантиметры", "см²" },
            { "square centimeters", "см²" },
            { "м²", "м²" },
            { "m²", "м²" },
            { "m2", "м²" },
            { "кв. м", "м²" },
            { "квадратные метры", "м²" },
            { "square meters", "м²" },
            { "км²", "км²" },
            { "km²", "км²" },
            { "km2", "км²" },
            { "кв. км", "км²" },
            { "квадратные километры", "км²" },
            { "square kilometers", "км²" },
            { "мм³", "мм³" },
            { "mm³", "мм³" },
            { "mm3", "мм³" },
            { "куб. мм", "мм³" },
            { "кубические миллиметры", "мм³" },
            { "cubic millimeters", "мм³" },
            { "см³", "см³" },
            { "cm³", "см³" },
            { "cm3", "см³" },
            { "куб. см", "см³" },
            { "кубические сантиметры", "см³" },
            { "cubic centimeters", "см³" },
            { "м³", "м³" },
            { "m³", "м³" },
            { "m3", "м³" },
            { "куб. м", "м³" },
            { "кубические метры", "м³" },
            { "cubic meters", "м³" },
            { "л", "л" },
            { "литр", "л" },
            { "литры", "л" },
            { "литров", "л" },
            { "l", "л" },
            { "кг", "кг" },
            { "килограмм", "кг" },
            { "килограммы", "кг" },
            { "килограммов", "кг" },
            { "kg", "кг" },
            { "т", "т" },
            { "тонна", "т" },
            { "тонны", "т" },
            { "тонн", "т" },
            { "н", "Н" },
            { "ньютон", "Н" },
            { "ньютоны", "Н" },
            { "ньютонов", "Н" },
            { "па", "Па" },
            { "паскаль", "Па" },
            { "паскали", "Па" },
            { "паскалей", "Па" },
            { "кпа", "кПа" },
            { "килопаскаль", "кПа" },
            { "килопаскали", "кПа" },
            { "килопаскалей", "кПа" },
            { "мпа", "МПа" },
            { "мегапаскаль", "МПа" },
            { "мегапаскали", "МПа" },
            { "мегапаскалей", "МПа" },
            { "вт", "Вт" },
            { "ватт", "Вт" },
            { "ватты", "Вт" },
            { "ваттов", "Вт" },
            { "квт", "кВт" },
            { "киловатт", "кВт" },
            { "киловатты", "кВт" },
            { "киловаттов", "кВт" },
            { "в", "В" },
            { "вольт", "В" },
            { "вольты", "В" },
            { "вольтов", "В" },
            { "а", "А" },
            { "ампер", "А" },
            { "амперы", "А" },
            { "амперов", "А" }
        };

        public SumParametersM(Document doc)
        {
            Doc = doc;
            ReloadFilteredView();
        }

        public Document Doc { get; private set; }

        public ObservableCollection<SumParameterResultM> SumResults { get; } = new ObservableCollection<SumParameterResultM>();

        public ICollectionView FilteredSumResults => _filteredSumResults;

        public Element[] UserSelElems
        {
            get => _userSelElems;
            private set => _userSelElems = value ?? new Element[0];
        }

        public int RoundDigits
        {
            get => _roundDigits;
            set
            {
                int normalizedValue = Math.Max(0, Math.Min(8, value));
                if (_roundDigits == normalizedValue)
                    return;

                _roundDigits = normalizedValue;
                NotifyPropertyChanged();
                RefreshCalculatedText();
            }
        }

        public string CoefficientText
        {
            get => _coefficientText;
            set
            {
                _coefficientText = value;
                NotifyPropertyChanged();
                RefreshCalculatedText();
            }
        }

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (value == null || value.StartsWith("KPLN_Parameters_Ribbon"))
                    return;

                _searchText = value;
                NotifyPropertyChanged();
                _filteredSumResults?.Refresh();
            }
        }

        public void SetUserSelection(Document doc, IEnumerable<Element> userSelElems)
        {
            Doc = doc;
            UserSelElems = GetUniqueElements(userSelElems);
            ReloadResults();
        }

        public static Element[] GetUserSelection(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp?.ActiveUIDocument;
            if (uidoc == null)
                return new Element[0];

            Document doc = uidoc.Document;
            return uidoc.Selection
                .GetElementIds()
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToArray();
        }

        public void IncrementRoundDigits() => RoundDigits++;

        public void DecrementRoundDigits() => RoundDigits--;

        public void ReloadResults()
        {
            SumResults.Clear();

            Dictionary<string, SumAccumulator> resultMap = new Dictionary<string, SumAccumulator>();
            foreach (Element element in UserSelElems.Where(IsSupportedElement))
                AddElementParametersToMap(element, resultMap);

            foreach (SumAccumulator accumulator in resultMap.Values
                .Where(r => r.ValueCount > 0)
                .OrderBy(r => r.ParameterName)
                .ThenBy(r => r.UnitName))
            {
                SumResults.Add(accumulator.ToResult());
            }

            RefreshCalculatedText();
            ReloadFilteredView();
        }

        public string GetTsv()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("Имя параметра\tЗначение\tЕд.изм.\tС коэффициентом\tБез коэффициента");

            foreach (SumParameterResultM result in SumResults)
            {
                builder.AppendLine(string.Join("\t", new[]
                {
                    result.ParameterName,
                    result.ValueText,
                    result.UnitName,
                    result.ValueWithCoefficientText,
                    result.ValueWithoutCoefficientText
                }));
            }

            return builder.ToString();
        }

        private Element[] GetUniqueElements(IEnumerable<Element> elements)
        {
            if (elements == null)
                return new Element[0];

            Dictionary<ElementId, Element> result = new Dictionary<ElementId, Element>();
            foreach (Element element in elements)
            {
                if (element == null || result.ContainsKey(element.Id))
                    continue;

                result.Add(element.Id, element);
            }

            return result.Values.ToArray();
        }

        private bool IsSupportedElement(Element element)
        {
            if (element?.Category == null)
                return false;

            if (element.Category.Name.Contains(".dwg"))
                return false;

            return element.Category.CategoryType == CategoryType.Model
                   || element.Category.CategoryType == CategoryType.Internal;
        }

        private void AddElementParametersToMap(Element element, Dictionary<string, SumAccumulator> resultMap)
        {
            AddParameterSetToMap(element.Parameters, resultMap);

            ElementId typeId = element.GetTypeId();
            if (Doc == null || typeId.Equals(ElementId.InvalidElementId))
                return;

            Element typeElement = Doc.GetElement(typeId);
            if (typeElement != null)
                AddParameterSetToMap(typeElement.Parameters, resultMap);
        }

        private void AddParameterSetToMap(ParameterSet parameters, Dictionary<string, SumAccumulator> resultMap)
        {
            foreach (Parameter parameter in parameters)
            {
                if (parameter?.Definition == null)
                    continue;

                if (parameter.StorageType == StorageType.ElementId || parameter.StorageType == StorageType.None)
                    continue;

                if (!TryGetDoubleValue(parameter, out double value))
                    continue;

                string parameterName = parameter.Definition.Name;
                string parameterIdText = GetParameterIdText(parameter);
                string displayParameterName = string.IsNullOrWhiteSpace(parameterIdText)
                    ? parameterName
                    : $"{parameterName} (id: {parameterIdText})";
                string unitName = GetUnitName(parameter, value);
                string key = $"{parameterName}\t{parameterIdText}\t{unitName}";
                double displayValue = GetProjectUnitValue(parameter, value);

                if (!resultMap.TryGetValue(key, out SumAccumulator accumulator))
                {
                    accumulator = new SumAccumulator(displayParameterName, parameterName, parameterIdText, unitName);
                    resultMap.Add(key, accumulator);
                }

                accumulator.AddValue(displayValue);
            }
        }

        private string GetParameterIdText(Parameter parameter)
        {
            try
            {
                return parameter.Id.IntegerValue.ToString(CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private bool TryGetDoubleValue(Parameter parameter, out double value)
        {
            value = 0;

            if (!parameter.HasValue)
                return false;

            try
            {
                switch (parameter.StorageType)
                {
                    case StorageType.Double:
                        value = parameter.AsDouble();
                        return true;
                    case StorageType.Integer:
                        value = parameter.AsInteger();
                        return true;
                    case StorageType.String:
                        return TryGetDoubleFromText(parameter.AsString(), out value)
                               || TryGetDoubleFromText(parameter.AsValueString(), out value);
                    default:
                        return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool TryGetDoubleFromText(string text, out double value)
        {
            value = 0;

            if (string.IsNullOrWhiteSpace(text))
                return false;

            string trimmedText = text.Trim();
            if (TryParseDouble(trimmedText, out value))
                return true;

            Match match = Regex.Match(trimmedText, @"^[-+]?\d+(?:[\s\u00A0]\d{3})*(?:[,.]\d+)?");
            if (!match.Success)
                return false;

            string numericPart = match.Value.Replace(" ", string.Empty).Replace("\u00A0", string.Empty);
            return TryParseDouble(numericPart, out value);
        }

        private bool TryParseDouble(string text, out double value)
        {
            value = 0;

            if (string.IsNullOrWhiteSpace(text))
                return false;

            NumberStyles styles = NumberStyles.Float | NumberStyles.AllowThousands;
            CultureInfo[] cultures =
            {
                CultureInfo.CurrentCulture,
                CultureInfo.InvariantCulture,
                new CultureInfo("ru-RU")
            };

            foreach (CultureInfo culture in cultures)
            {
                if (double.TryParse(text, styles, culture, out value))
                    return true;
            }

            string dotText = text.Replace(',', '.');
            return double.TryParse(dotText, styles, CultureInfo.InvariantCulture, out value);
        }

        private double GetProjectUnitValue(Parameter parameter, double internalValue)
        {
            if (parameter.StorageType != StorageType.Double)
                return internalValue;

            try
            {
#if Debug2020 || Revit2020
                DisplayUnitType displayUnit = parameter.DisplayUnitType;
                return UnitUtils.ConvertFromInternalUnits(internalValue, displayUnit);
#else
                ForgeTypeId unitTypeId = parameter.GetUnitTypeId();
                return UnitUtils.ConvertFromInternalUnits(internalValue, unitTypeId);
#endif
            }
            catch (Exception)
            {
                return internalValue;
            }
        }

        private string GetUnitName(Parameter parameter, double internalValue)
        {
            if (parameter.StorageType != StorageType.Double)
                return string.Empty;

            try
            {
                string valueString = GetFormattedProjectValue(parameter, internalValue);
                if (string.IsNullOrWhiteSpace(valueString))
                    return GetFallbackUnitName(parameter);

                Match match = Regex.Match(valueString.Trim(), @"^[-+]?\d+(?:[\s\u00A0]\d{3})*(?:[,.]\d+)?\s*(.*)$");
                if (!match.Success)
                    return GetFallbackUnitName(parameter);

                string unitName = NormalizeUnitName(match.Groups[1].Value);
                return string.IsNullOrWhiteSpace(unitName)
                    ? GetFallbackUnitName(parameter)
                    : unitName;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private string GetFormattedProjectValue(Parameter parameter, double internalValue)
        {
            if (Doc == null)
                return string.Empty;

            try
            {
#if Debug2020 || Revit2020
                UnitType unitType = parameter.Definition.UnitType;
                return UnitFormatUtils.Format(Doc.GetUnits(), unitType, internalValue, false, false);
#else
                ForgeTypeId spec = parameter.Definition.GetDataType();
                if (!UnitUtils.IsMeasurableSpec(spec))
                    return string.Empty;

                return UnitFormatUtils.Format(Doc.GetUnits(), spec, internalValue, forEditing: false);
#endif
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private string GetFallbackUnitName(Parameter parameter)
        {
            try
            {
#if Debug2020 || Revit2020
                return NormalizeUnitName(LabelUtils.GetLabelFor(parameter.DisplayUnitType));
#else
                return NormalizeUnitName(LabelUtils.GetLabelForUnit(parameter.GetUnitTypeId()));
#endif
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        private string NormalizeUnitName(string unitName)
        {
            if (string.IsNullOrWhiteSpace(unitName))
                return string.Empty;

            string normalizedUnitName = Regex.Replace(unitName.Trim().Trim('.', ',', ';', ':', '(', ')'), @"\s+", " ");
            return UnitNameAbbreviations.TryGetValue(normalizedUnitName, out string abbreviation)
                ? abbreviation
                : normalizedUnitName;
        }

        private void RefreshCalculatedText()
        {
            double coefficient = GetCoefficient();
            foreach (SumParameterResultM result in SumResults)
            {
                result.ValueText = FormatValue(result.Sum);
                result.ValueWithCoefficientText = FormatValue(result.Sum * (1 + coefficient));
                result.ValueWithoutCoefficientText = FormatValue(result.Sum);
                result.NotifyAll();
            }
        }

        private double GetCoefficient()
        {
            string coefficientText = CoefficientText?.Replace("%", string.Empty);
            if (TryParseDouble(coefficientText, out double coefficientPercent))
                return coefficientPercent / 100;

            return 0;
        }

        private string FormatValue(double value) =>
            Math.Round(value, RoundDigits).ToString($"F{RoundDigits}", CultureInfo.CurrentCulture);

        private void ReloadFilteredView()
        {
            _filteredSumResults = CollectionViewSource.GetDefaultView(SumResults);
            _filteredSumResults.Filter = FilterMethod;
            NotifyPropertyChanged(nameof(FilteredSumResults));
        }

        private bool FilterMethod(object obj)
        {
            if (string.IsNullOrWhiteSpace(SearchText))
                return true;

            if (!(obj is SumParameterResultM result) || string.IsNullOrWhiteSpace(result.ParameterName))
                return false;

            string parameterName = string.IsNullOrWhiteSpace(result.SearchParameterName)
                ? result.ParameterName
                : result.SearchParameterName;

            return parameterName.ToLower().Contains(SearchText.ToLower());
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private sealed class SumAccumulator
        {
            public SumAccumulator(string parameterName, string searchParameterName, string parameterIdText, string unitName)
            {
                ParameterName = parameterName;
                SearchParameterName = searchParameterName;
                ParameterIdText = parameterIdText;
                UnitName = unitName;
            }

            public string ParameterName { get; }

            public string SearchParameterName { get; }

            public string ParameterIdText { get; }

            public string UnitName { get; }

            public int ValueCount { get; private set; }

            public double Sum { get; private set; }

            public void AddValue(double value)
            {
                Sum += value;
                ValueCount++;
            }

            public SumParameterResultM ToResult()
            {
                return new SumParameterResultM
                {
                    ParameterName = ParameterName,
                    SearchParameterName = SearchParameterName,
                    UnitName = UnitName,
                    Sum = Sum
                };
            }
        }
    }
}
