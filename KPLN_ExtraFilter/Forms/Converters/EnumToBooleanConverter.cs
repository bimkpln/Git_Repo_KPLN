using System;
using System.Globalization;
using System.Windows.Data;

namespace KPLN_ExtraFilter.Forms.Converters
{
    public sealed class EnumToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value?.ToString() == parameter?.ToString();

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => (bool)value ? Enum.Parse(targetType, parameter.ToString()) : Binding.DoNothing;
    }
}
