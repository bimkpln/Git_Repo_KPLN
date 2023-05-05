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
    public class WPFEntityApprCommentVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            WPFEntity entity = value as WPFEntity;
            if (entity != null)
            {
                return entity.ApproveComment == null ? Visibility.Hidden : Visibility.Visible;
            }

            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}