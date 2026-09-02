using KPLN_MEPSpacing.Forms.Entities;
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace KPLN_MEPSpacing.Forms.Converters
{
    public sealed class UserMessageLevelToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is UserMessageLevel level))
                level = UserMessageLevel.Neutral;

            string target = parameter?.ToString();
            if (target == "Border")
                return GetBorderBrush(level);

            return GetBackgroundBrush(level);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) => Binding.DoNothing;

        private static Brush GetBackgroundBrush(UserMessageLevel level)
        {
            switch (level)
            {
                case UserMessageLevel.Success:
                    return new SolidColorBrush(Color.FromRgb(31, 85, 57));
                case UserMessageLevel.Warning:
                    return new SolidColorBrush(Color.FromRgb(116, 72, 27));
                case UserMessageLevel.Error:
                    return new SolidColorBrush(Color.FromRgb(105, 42, 45));
                default:
                    return new SolidColorBrush(Color.FromRgb(51, 56, 64));
            }
        }

        private static Brush GetBorderBrush(UserMessageLevel level)
        {
            switch (level)
            {
                case UserMessageLevel.Success:
                    return new SolidColorBrush(Color.FromRgb(72, 186, 119));
                case UserMessageLevel.Warning:
                    return new SolidColorBrush(Color.FromRgb(232, 151, 55));
                case UserMessageLevel.Error:
                    return new SolidColorBrush(Color.FromRgb(235, 90, 94));
                default:
                    return new SolidColorBrush(Color.FromRgb(80, 86, 96));
            }
        }
    }
}
