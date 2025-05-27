using Autodesk.Revit.DB;
using System;
using System.Windows;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Input;
using System.Windows.Controls;
using System.Globalization;
using System.Windows.Media;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace KPLN_Tools.Forms
{
    public partial class AR_TEPDesign_paramNameSelect : Window
    {
        Document _doc;
        List<ElementId> _elementIds = new List<ElementId>();

        int countUniqueValues;

        private List<string> AllLocations = new List<string>
        {
            "Не требуется",
            "Сверху слева",
            "Сверху справа",
            "Снизу слева",
            "Снизу справа"
        };

        private System.Windows.Media.Color selectedColor = System.Windows.Media.Colors.LightGray;
        double fontSizeInDoc;

        public string SelectedParamName { get; private set; }

        public string SelectedEmptyLocation
        {
            get
            {
                return ComboBoxEmptyLocation.SelectedItem as string;
            }
        }

        public string SelectedTableSortType
        {
            get
            {
                return (ComboBoxTableSortType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
        }

        public System.Windows.Media.Color SelectedDefaultColor
        {
            get
            {
                return selectedColor;
            }
        }

        public string SelectedColorBindingType
        {
            get
            {
                return (ComboBoxColorBinding.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
        }

        public double? SelectedLightenFactor
        {
            get
            {
                if (double.TryParse(TextBoxLighten.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) return val;
                return null;
            }
        }

        public string SelectedFontName
        {
            get
            {
                return ComboBoxFontName.SelectedItem as string;
            }
        }

        public double? SelectedFontSize
        {
            get
            {
                if (double.TryParse(TextBoxFontSize.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) return val;
                return null;
            }
        }

        public AR_TEPDesign_paramNameSelect(Document doc, List<ElementId> elementIds)
        {
            InitializeComponent();
            _doc = doc;
            _elementIds = elementIds;

            if (elementIds == null || elementIds.Count == 0 || doc == null)
                return;

            var firstElement= doc.GetElement(elementIds.First());
            var candidateParams = firstElement.Parameters
                .Cast<Parameter>()
                .Where(p => p.StorageType == StorageType.String)
                .Select(p => p.Definition.Name)
                .ToList();

            foreach (var id in elementIds.Skip(1))
            {
                var el = doc.GetElement(id);
                var currentParamNames = el.Parameters
                    .Cast<Parameter>()
                    .Where(p => p.StorageType == StorageType.String)
                    .Select(p => p.Definition.Name)
                    .ToHashSet();

                candidateParams = candidateParams.Where(name => currentParamNames.Contains(name)).ToList();
            }

            ComboBoxParams.ItemsSource = candidateParams.OrderBy(name => name).ToList();

            if (candidateParams.Contains("Имя"))
            {
                ComboBoxParams.SelectedItem = "Имя";
            }
            else if (candidateParams.Contains("Назначение"))
            {
                ComboBoxParams.SelectedItem = "Назначение";
            }

            ComboBoxEmptyLocation.SelectedIndex = 0; 
            ComboBoxEmptyLocation.IsEnabled = false;

            countUniqueValues = GetUniqueParamValuesCount(doc, elementIds, ComboBoxParams.SelectedItem as string);             
 
            if (countUniqueValues == 0)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsCount.Text = "Параметр не имеет уникальных значений";
                TextBlockParamsTableInfo.Text = "Невозможно сформировать таблицу";
            }
            else if (countUniqueValues == 1)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - 1;";
                TextBlockParamsTableInfo.Text = "Таблица с одним столбцом. Вы уверены?";
            }
            else if (countUniqueValues <= 12)
            {
                if (countUniqueValues % 2 == 0)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {countUniqueValues / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsTableInfo.Text = "Таблица будет сформирована без дополнительных условий";
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                    TextBlockParamsTableInfo.Text = $"Для формирования таблицы требуется ячейка-заглушка";
                }
            }
            else 
            {              
                if (countUniqueValues % 2 == 0)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {countUniqueValues / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                    TextBlockParamsTableInfo.Text = $"Слишком большое кол-во столбцов, таблица может быть сформирована с ошибками";
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                    TextBlockParamsTableInfo.Text = $"Слишком большое кол-во столбцов. Для формирования таблицы требуется ячейка-заглушка";
                }
            }

            UpdateComboBoxEmptyLocation(countUniqueValues);

            var textTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .ToList();

            var fontNames = textTypes
                .Select(t => t.get_Parameter(BuiltInParameter.TEXT_FONT)?.AsString())
                .Where(f => !string.IsNullOrEmpty(f))
                .Distinct()
                .OrderBy(f => f)
                .ToList();

            ComboBoxFontName.ItemsSource = fontNames;

            string defaultFont = "GOST Common";
            if (fontNames.Contains(defaultFont))
                ComboBoxFontName.SelectedItem = defaultFont;
            else
                ComboBoxFontName.SelectedIndex = 0;

            double? fontSizeFeet = textTypes.FirstOrDefault()?.get_Parameter(BuiltInParameter.TEXT_SIZE)?.AsDouble();

            if (fontSizeFeet.HasValue)
            {
                fontSizeInDoc = fontSizeFeet.Value * 304.8;
                TextBoxFontSize.Text = fontSizeInDoc.ToString("0.##"); 
            }
        }

        public int GetUniqueParamValuesCount(Document doc, List<ElementId> elementIds, string paramName)
        {
            HashSet<string> uniqueValues = new HashSet<string>();

            foreach (ElementId id in elementIds)
            {
                Element element = doc.GetElement(id);
                if (element == null) continue;

                Parameter param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    string value = "";

                    switch (param.StorageType)
                    {
                        case StorageType.Double:
                            value = param.AsValueString() ?? param.AsDouble().ToString();
                            break;
                        case StorageType.Integer:
                            value = param.AsValueString() ?? param.AsInteger().ToString();
                            break;
                        case StorageType.String:
                            value = param.AsString();
                            break;
                        case StorageType.ElementId:
                            ElementId elemId = param.AsElementId();
                            value = elemId.IntegerValue >= 0
                                ? doc.GetElement(elemId)?.Name ?? elemId.IntegerValue.ToString()
                                : elemId.IntegerValue.ToString();
                            break;
                        default:
                            continue;
                    }

                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        uniqueValues.Add(value);
                    }
                }
            }

            return uniqueValues.Count;
        }

        private void UpdateComboBoxEmptyLocation(int countUniqueValues)
        {
            if (countUniqueValues == 0)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsCount.Text = "Параметр не имеет уникальных значений";
                TextBlockParamsTableInfo.Text = "Невозможно сформировать таблицу";

                ComboBoxEmptyLocation.ItemsSource = AllLocations;
                ComboBoxEmptyLocation.SelectedItem = "Не требуется";
                ComboBoxEmptyLocation.IsEnabled = false;
            }
            else if (countUniqueValues == 1)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - 1;";
                TextBlockParamsTableInfo.Text = "Таблица с одним столбцом. Вы уверены?";

                ComboBoxEmptyLocation.ItemsSource = AllLocations;
                ComboBoxEmptyLocation.SelectedItem = "Не требуется";
                ComboBoxEmptyLocation.IsEnabled = false;
            }
            else if (countUniqueValues <= 12)
            {
                if (countUniqueValues % 2 == 0)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {countUniqueValues / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsTableInfo.Text = "Таблица будет сформирована без дополнительных условий";

                    ComboBoxEmptyLocation.ItemsSource = AllLocations;
                    ComboBoxEmptyLocation.SelectedItem = "Не требуется";
                    ComboBoxEmptyLocation.IsEnabled = false;
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                    TextBlockParamsTableInfo.Text = $"Для формирования таблицы требуется ячейка-заглушка";

                    var listWithoutDefault = AllLocations.Where(x => x != "Не требуется").ToList();
                    ComboBoxEmptyLocation.ItemsSource = listWithoutDefault;
                    ComboBoxEmptyLocation.IsEnabled = true;
                    ComboBoxEmptyLocation.SelectedItem = "Снизу справа";
                }
            }
            else
            {
                if (countUniqueValues % 2 == 0)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.MediumVioletRed);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {countUniqueValues / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.MediumVioletRed);
                    TextBlockParamsTableInfo.Text = $"Слишком большое кол-во столбцов, таблица может быть сформирована с ошибками";

                    ComboBoxEmptyLocation.ItemsSource = AllLocations;
                    ComboBoxEmptyLocation.SelectedItem = "Не требуется";
                    ComboBoxEmptyLocation.IsEnabled = false;
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.MediumVioletRed);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                    TextBlockParamsTableInfo.Text = $"Слишком большое кол-во столбцов. Для формирования таблицы требуется ячейка-заглушка";

                    var listWithoutDefault = AllLocations.Where(x => x != "Не требуется").ToList();
                    ComboBoxEmptyLocation.ItemsSource = listWithoutDefault;
                    ComboBoxEmptyLocation.IsEnabled = true;
                    ComboBoxEmptyLocation.SelectedItem = "Снизу справа";
                }
            }
        }

        private void ComboBoxParams_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboBoxParams.SelectedItem == null) return;

            countUniqueValues = GetUniqueParamValuesCount(_doc, _elementIds, ComboBoxParams.SelectedItem as string);
            TextBlockParamsCount.Text = $"Найдено уникальных значений параметра: {countUniqueValues}";

            if (countUniqueValues == 0 || countUniqueValues > 12)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                if (countUniqueValues > 12 && !(countUniqueValues % 2 == 0))
                {
                    TextBlockParamsCount.Text = $"Найдено уникальных значений параметра: {countUniqueValues} (требуется заглушка)";
                }
            }
            else
            {
                if (countUniqueValues % 2 == 0 || countUniqueValues == 1)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Blue);
                    TextBlockParamsCount.Text = $"Найдено уникальных значений параметра: {countUniqueValues} (требуется заглушка)";
                }
            }

            UpdateComboBoxEmptyLocation(countUniqueValues);
        }

        private void ColorPickerButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new ColorDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var drawingColor = dialog.Color;
                    selectedColor = System.Windows.Media.Color.FromArgb(
                        drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);


                    ColorPickerButton.Background = new SolidColorBrush(selectedColor);
                    ColorPickerButton.Content = selectedColor.ToString();
                }
            }
        }

        private void NumberOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            string fullText = textBox.Text.Insert(textBox.SelectionStart, e.Text);

            bool hasDotOrComma = fullText.Count(c => c == '.' || c == ',') <= 1;

            e.Handled = !(double.TryParse(fullText.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _) && hasDotOrComma);
        }

          
        private void TextBoxLighten_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(
                TextBoxLighten.Text.Replace(',', '.'),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value))
            {
                value = Math.Max(0.0, Math.Min(1.0, value));
                TextBoxLighten.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxLighten.Text = "0.5";
            }
        }

        private void TextBoxFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(
                TextBoxFontSize.Text.Replace(',', '.'),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value))
            {
                value = Math.Max(0.5, Math.Min(9.0, value));
                TextBoxFontSize.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxFontSize.Text = fontSizeInDoc.ToString("0.##");
            }
        }


        private void ButtonOk_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxParams.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Выберите параметр для формирования таблицы.", "Предупреждение");
                return;
            }

            if (ComboBoxEmptyLocation.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Выберите месторасположение заглушки.", "Предупреждение");
                return;
            }

            if (ComboBoxTableSortType.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Выберите тип сортировки таблицы.", "Предупреждение");
                return;
            }

            if (string.IsNullOrWhiteSpace(TextBoxLighten.Text))
            {
                System.Windows.MessageBox.Show("Значение `Коэффициент осветления` не указано. Установлено значение по умолчанию 0.5. Если необходимо изменить — введите новое значение и повторите.", "Предупреждение");
                TextBoxLighten.Text = "0.5";
                return;
            }

            if (ComboBoxFontName.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Выберите имя шрифта.", "Предупреждение");
                return;
            }

            if (string.IsNullOrWhiteSpace(TextBoxFontSize.Text))
            {
                System.Windows.MessageBox.Show("Размер шрифта не указан. Введите значение от 0.5 до 9.", "Предупреждение");
                TextBoxFontSize.Text = fontSizeInDoc.ToString("0.##");
                return;
            }

            if (countUniqueValues == 0)
            {
                System.Windows.MessageBox.Show("Выбран параметр, который не содержит значений.\nВыберите другой параметр и повторите попытку.", "Ошибка");
                return;
            }

            SelectedParamName = ComboBoxParams.SelectedItem.ToString();

            DialogResult = true;
            Close();
        }
    }
}
