using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Класс для переопределения видимости полей WPFEntity в xml
    /// </summary>
    public class StringToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null) return Visibility.Collapsed;
            
            if (string.IsNullOrEmpty((string)value)) return Visibility.Collapsed;
            else return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}