using System;
using System.Globalization;
using System.Windows.Data;

namespace KPLN_DefaultPanelExtension_Modify.Forms.ValueConvertors
{
    /// <summary>
    /// Для привязки IsChecked к enum SelectedAlign (toggle как радио)
    /// </summary>
    internal sealed class EnumEqualsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null) return false;
            return string.Equals(value.ToString(), parameter.ToString(), StringComparison.OrdinalIgnoreCase);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Не используем обратное преобразование, выбор идёт через Command
            return Binding.DoNothing;
        }
    }
}
