using System;
using System.Globalization;
using System.Windows.Data;

namespace KPLN_ModelChecker_Batch.Forms.Converters
{
    public class TextToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null)
                return false;

            return value.ToString().Equals(parameter.ToString(), StringComparison.OrdinalIgnoreCase);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolean && boolean)
                return parameter.ToString();

            return Binding.DoNothing;
        }
    }
}
