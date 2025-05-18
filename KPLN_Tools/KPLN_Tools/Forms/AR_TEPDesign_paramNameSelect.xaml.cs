using Autodesk.Revit.DB;
using System;
using System.Windows;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Input;
using System.Windows.Controls;
using System.Globalization;

namespace KPLN_Tools.Forms
{
    public partial class AR_TEPDesign_paramNameSelect : Window
    {
        public string SelectedParamName { get; private set; }

        public double? SelectedTableHeight
        {
            get
            {
                if (double.TryParse(TextBoxTableHeight.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double height)) return height / 304.8;
                return null;
            }
        }

        public double? SelectedFontSize
        {
            get
            {
                if (double.TryParse(TextBoxFontSize.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double size)) return size / 304.8;
                return null;
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

        public double? SelectedDarkenFactor
        {
            get
            {
                if (double.TryParse( TextBoxDarken.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) return val;              
                return null;
            }
        }

        double _fontSize;

        public AR_TEPDesign_paramNameSelect(Document doc, List<ElementId> elementIds, double fontSize)
        {
            InitializeComponent();

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

            _fontSize = fontSize;
            TextBoxTableHeight.Text = "90.00";           
            TextBoxFontSize.Text = _fontSize.ToString();
            TextBoxLighten.Text = "0.6";
            TextBoxDarken.Text = "0.1";
        }

        private void NumberOnly_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            string fullText = textBox.Text.Insert(textBox.SelectionStart, e.Text);

            bool hasDotOrComma = fullText.Count(c => c == '.' || c == ',') <= 1;

            e.Handled = !(double.TryParse(fullText.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _) && hasDotOrComma);
        }

        private void TextBoxTableHeight_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxTableHeight.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                value = Math.Max(30, Math.Min(200, value));
                TextBoxTableHeight.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxTableHeight.Text = "";
            }
        }

        private void TextBoxFontSize_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(TextBoxFontSize.Text.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            {
                value = Math.Max(0.5, Math.Min(100, value));
                TextBoxFontSize.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxFontSize.Text = "";
            }
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
                TextBoxLighten.Text = "0";
            }
        }

        private void TextBoxDarken_LostFocus(object sender, RoutedEventArgs e)
        {
            if (double.TryParse(
                TextBoxDarken.Text.Replace(',', '.'),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double value))
            {
                value = Math.Max(0.0, Math.Min(1.0, value));
                TextBoxDarken.Text = value.ToString("0.##", CultureInfo.InvariantCulture);
            }
            else
            {
                TextBoxDarken.Text = "0";
            }
        }

        private void ButtonOk_Click(object sender, RoutedEventArgs e)
        {
            if (ComboBoxParams.SelectedItem == null)
            {
                MessageBox.Show("Выберите параметр для формирования таблицы.", "Предупреждение");
                return;
            }

            if (TextBoxTableHeight.Text == "" || TextBoxTableHeight.Text == null)
            {
                MessageBox.Show("Значение `Высота ячейки (мм)` не было указано. Было возвращено стандартное значение. Если вам необходимо поменять значение - укажите его и повторите попытку.", "Предупреждение");
                TextBoxTableHeight.Text = "90.00";
                return;
            }

            if (TextBoxFontSize.Text == "" || TextBoxFontSize.Text == null)
            {
                MessageBox.Show("Значение `Размер шрифта (мм)` не было указано. Было возвращено стандартное значение. Если вам необходимо поменять значение - укажите его и повторите попытку.", "Предупреждение");
                TextBoxTableHeight.Text = _fontSize.ToString();
                return;
            }

            if (TextBoxLighten.Text == "" || TextBoxLighten.Text == null)
            {
                MessageBox.Show("Значение `Коэффициент осветления` не было указано. Было возвращено стандартное значение. Если вам необходимо поменять значение - укажите его и повторите попытку.", "Предупреждение");
                TextBoxLighten.Text = "0.6";
                return;
            }

            if (TextBoxDarken.Text == "" || TextBoxDarken.Text == null)
            {
                MessageBox.Show("Значение `Коэффициент затемнения` не было указано. Было возвращено стандартное значение. Если вам необходимо поменять значение - укажите его и повторите попытку.", "Предупреждение");
                TextBoxDarken.Text = "0.1";
                return;
            }

            SelectedParamName = ComboBoxParams.SelectedItem.ToString();

            DialogResult = true;
            Close();
        }
    }
}
