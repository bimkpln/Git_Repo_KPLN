using System;
using System.Collections.Generic;
using System.Windows.Data;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Конвертор преобразования массива текста в строки wpf
    /// </summary>
    public class JoinStringsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var strings = value as IEnumerable<string>;
            return string.Join(Environment.NewLine, strings);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
