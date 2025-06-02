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

namespace KPLN_Tools.Forms
{
    public partial class AR_TEPDesign_paramNameSelect : Window
    {
        Document _doc;

        List<ElementId> _elementIds = new List<ElementId>();
        int countUniqueValues;
        int suggestedRowCount;

        private System.Windows.Media.Color selectedColorEmptyColorScheme = System.Windows.Media.Colors.LightGray; // Значение при отсутствии цветовой схемы
        private System.Windows.Media.Color selectedColorDummyСell = System.Windows.Media.Colors.LightGray; // Значение заглушки

        public string SelectedParamName { get; private set; } // Формирующий параметр

        // Компановка. Кол-во столбцов
        public int SelectedRowCount
        {
            get
            {
                if (int.TryParse(TextBoxRowCount.Text, out int result))
                    return result;
                else
                    return 0;
            }
        }

        // Компановка. Место расположение "заглушки"
        public string SelectedEmptyLocation
        {
            get
            {
                return (ComboBoxEmptyLocation.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
        }

        // Компановка. Тип сортировки данных
        public string SelectedTableSortType
        {
            get
            {
                return (ComboBoxTableSortType.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
        }

        // Цвет. Значение при отсутствии цветовой схемы
        public System.Windows.Media.Color SelectedColorEmptyColorScheme
        {
            get
            {
                return selectedColorEmptyColorScheme;
            }
        }
    
        // Цвет. Цвет заглушки
        public System.Windows.Media.Color SelectedColorDummyСell
        {
            get
            {
                return selectedColorDummyСell;
            }
        }

        // Цвет. Приоритет срабатывания заглушки
        public bool SelectedELPriority
        {
            get
            {
                return ColorPickerButtonELPriority.IsChecked == true;
            }
        }

        // Цвет. Приоритет окрашивания ячеек
        public string SelectedColorBindingType
        {
            get
            {
                return (ComboBoxColorBinding.SelectedItem as ComboBoxItem)?.Content?.ToString();
            }
        }

        // Цвет. Коэф. осветления всех ячеек
        public double? SelectedLightenFactor
        {
            get
            {
                if (double.TryParse(TextBoxLighten.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) return val;
                return null;
            }
        }

        // Цвет. Коэф. осветления ячеек по отношению первого ряда ко второму
        public double? SelectedLightenFactorRow
        {
            get
            {
                if (double.TryParse(TextBoxLightenRow.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) return val;
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


            countUniqueValues = GetUniqueParamValuesCount(doc, elementIds, ComboBoxParams.SelectedItem as string);
            UpdateComboBoxEmptyLocation(countUniqueValues);



            var fontNames = new FilteredElementCollector(_doc)
                .OfClass(typeof(TextNoteType))
                .Cast<TextNoteType>()
                .Select(t => t.get_Parameter(BuiltInParameter.TEXT_FONT)?.AsString())
                .Where(name => !string.IsNullOrEmpty(name))
                .Distinct()
                .OrderBy(name => name)
                .ToList();
        }


        /// <summary>
        /// Получение кол-ва значений у уникального параметра
        /// </summary>
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

        /// <summary>
        /// Обновление данных
        /// </summary>
        private void UpdateComboBoxEmptyLocation(int countUniqueValues)
        {
            if (countUniqueValues == 0)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                TextBlockParamsCount.Text = "Параметр не имеет уникальных значений";
                TextBlockParamsTableInfo.Text = "Невозможно сформировать таблицу";

                TextBoxRowCount.Text = "0";
                TextBoxRowCount.IsEnabled = false;
            }
            else if (countUniqueValues == 1)
            {
                TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - 1;";
                TextBlockParamsTableInfo.Text = "Таблица с одним столбцом. Вы уверены?";

                TextBoxRowCount.Text = "1";
                TextBoxRowCount.IsEnabled = false;
            }
            else if (countUniqueValues <= 12)
            {
                if (countUniqueValues % 2 == 0)
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {countUniqueValues / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsTableInfo.Text = "Таблица будет сформирована без дополнительных условий";

                    TextBoxRowCount.IsEnabled = true;
                    suggestedRowCount = (countUniqueValues % 2 == 0) ? countUniqueValues / 2 : (countUniqueValues + 1) / 2;                   
                    TextBoxRowCount.Text = $"{suggestedRowCount}";                
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.Green);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Blue);
                    TextBlockParamsTableInfo.Text = $"Для формирования таблицы требуется ячейка-заглушка";

                    TextBoxRowCount.IsEnabled = true;
                    suggestedRowCount = (countUniqueValues % 2 == 0) ? countUniqueValues / 2 : (countUniqueValues + 1) / 2;
                    TextBoxRowCount.Text = $"{suggestedRowCount}";
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

                    TextBoxRowCount.IsEnabled = true;
                    suggestedRowCount = (countUniqueValues % 2 == 0) ? countUniqueValues / 2 : (countUniqueValues + 1) / 2;
                    TextBoxRowCount.Text = $"{suggestedRowCount}";
                }
                else
                {
                    TextBlockParamsCount.Foreground = new SolidColorBrush(Colors.MediumVioletRed);
                    TextBlockParamsCount.Text = $"Уникальные значения параметра: {countUniqueValues}. Столбцов - {(countUniqueValues + 1) / 2}";
                    TextBlockParamsTableInfo.Foreground = new SolidColorBrush(Colors.Red);
                    TextBlockParamsTableInfo.Text = $"Слишком большое кол-во столбцов. Для формирования таблицы требуется ячейка-заглушка";

                    TextBoxRowCount.IsEnabled = true;
                    suggestedRowCount = (countUniqueValues % 2 == 0) ? countUniqueValues / 2 : (countUniqueValues + 1) / 2;
                    TextBoxRowCount.Text = $"{suggestedRowCount}";
                }
            }
        }

        /// <summary>
        /// XAML. Обновление данных после выбора параметра для формирования ТЭП
        /// </summary>
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

        /// <summary>
        /// XAML. Переопределение пользовательского значения кол-ва колонок
        /// </summary>
        private void TextBoxRowCount_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !int.TryParse(e.Text, out _);
        }

        /// <summary>
        /// XAML. Установка нецелочисленых значений в параметры осветления
        /// </summary>
        private void NumberOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            System.Windows.Controls.TextBox textBox = sender as System.Windows.Controls.TextBox;
            string fullText = textBox.Text.Insert(textBox.SelectionStart, e.Text);

            bool hasDotOrComma = fullText.Count(c => c == '.' || c == ',') <= 1;

            e.Handled = !(double.TryParse(fullText.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _) && hasDotOrComma);
        }

        /// <summary>
        /// XAML. Цифровое значение в общем коэф. осветления
        /// </summary>
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

        /// <summary>
        /// XAML. Цифровое значение в коэф. осветления одной строки к другой
        /// </summary>
        private void TextBoxLightenRow_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(
                TextBoxLightenRow.Text.Replace(',', '.'),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value))
            {
                value = Math.Max(0.0, Math.Min(1.0, value));
                TextBoxLightenRow.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxLightenRow.Text = "0.5";
            }
        }

        /// <summary>
        /// XAML. Цвет параметров без цветовой схемы
        /// </summary>
        private void ColorPickerButton_ClickESC(object sender, RoutedEventArgs e)
        {
            using (var dialog = new ColorDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var drawingColor = dialog.Color;
                    selectedColorEmptyColorScheme = System.Windows.Media.Color.FromArgb(
                        drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);


                    ColorPickerButtonESC.Background = new SolidColorBrush(selectedColorEmptyColorScheme);
                    ColorPickerButtonESC.Content = selectedColorEmptyColorScheme.ToString();
                }
            }
        }

        /// <summary>
        /// XAML. Цвет ячеек-заглушек
        /// </summary>
        private void ColorPickerButton_ClickEL(object sender, RoutedEventArgs e)
        {
            using (var dialog = new ColorDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var drawingColor = dialog.Color;
                    selectedColorDummyСell = System.Windows.Media.Color.FromArgb(
                        drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B);


                    ColorPickerButtonEL.Background = new SolidColorBrush(selectedColorDummyСell);
                    ColorPickerButtonEL.Content = selectedColorDummyСell.ToString();
                }
            }
        }



        /// <summary>
        /// XAML. Кнопка OK
        /// </summary>
        private void ButtonOk_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxParams.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("Выберите параметр для формирования таблицы.", "Предупреждение");
                return;
            }

            int minAllowed = (countUniqueValues % 2 == 0) ? countUniqueValues / 2 : (countUniqueValues + 1) / 2;
            int rowCount = int.Parse(TextBoxRowCount.Text);
            if (rowCount < minAllowed || rowCount > countUniqueValues)
            {
                suggestedRowCount = minAllowed;
                System.Windows.MessageBox.Show($"Параметр 'Пользовательское переопрределение кол-ва столбцов' указан не верно. Необходимо указать значение от {minAllowed} до {countUniqueValues}","Ошибка");
                TextBoxRowCount.Text = $"{minAllowed}";
                return;
            }


            if (string.IsNullOrWhiteSpace(TextBoxLighten.Text))
            {
                System.Windows.MessageBox.Show("Значение `Коэффициент изменения цвета (общий)` не указано. Установлено значение по умолчанию 0.5. Если необходимо изменить — введите новое значение и повторите.", "Предупреждение");
                TextBoxLighten.Text = "0.5";
                return;
            }

            if (string.IsNullOrWhiteSpace(TextBoxLightenRow.Text))
            {
                System.Windows.MessageBox.Show("Значение `Коэффициент изменения цвета (строка к строке)` не указано. Установлено значение по умолчанию 0.5. Если необходимо изменить — введите новое значение и повторите.", "Предупреждение");
                TextBoxLightenRow.Text = "0.5";
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
