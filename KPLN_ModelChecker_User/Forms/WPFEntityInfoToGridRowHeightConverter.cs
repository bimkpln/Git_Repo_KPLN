using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Класс для переопределения видимости высоты строк для WPFEntity
    /// </summary>
    [ValueConversion(typeof(WPFEntity), typeof(GridLength))]
    public class WPFEntityInfoToGridRowHeightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            WPFEntity entity = value as WPFEntity;
            if (entity != null)
            {
                return (entity.Info == null) ? new GridLength(0) : new GridLength(1, GridUnitType.Star);
            }

            return new GridLength(0);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}