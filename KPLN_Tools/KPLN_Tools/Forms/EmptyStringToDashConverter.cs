using System;
using System.Globalization;
using System.Windows.Data;

namespace KPLN_Tools.Forms
{
    /// <summary>
    /// Конвертор пустого текстового значения
    /// </summary>
    public class EmptyStringToDashConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is string data) || data == string.Empty)
            {
                return "-";
            }

            return data;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
