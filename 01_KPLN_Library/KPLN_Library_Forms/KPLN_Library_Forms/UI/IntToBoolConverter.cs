using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace KPLN_Library_Forms.UI
{
    /// <summary>
    /// Класс для переопределения IsEnable кнопки по длине строки
    /// </summary>
    [ValueConversion(typeof(int), typeof(bool))]
    public class IntToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                return (System.Convert.ToInt32(value) > 3);
            }
            catch (InvalidCastException)
            {
                return DependencyProperty.UnsetValue;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return System.Convert.ToBoolean(value) ? 1 : 0;
        }
    }
}
