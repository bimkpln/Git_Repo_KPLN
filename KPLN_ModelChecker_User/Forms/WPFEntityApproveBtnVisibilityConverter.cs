using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Класс для переопределения видимости кнопки подтверждения для WPFEntity
    /// </summary>
    [ValueConversion(typeof(WPFEntity), typeof(GridLength))]
    public class WPFEntityApproveBtnVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is WPFEntity entity)
                return entity.CanApproved ? Visibility.Visible : Visibility.Hidden;

            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}