using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Класс для переопределения видимости данных в поле WPFEntity в xml
    /// </summary>
    public class ElementIdListToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is IEnumerable<ElementId> list)
            {
                if (list.Count() > 1)
                {
                    return string.Join(", ", list);
                }
                else
                {
                    return list.FirstOrDefault();
                }
            }
            else if (value !=null && int.TryParse(value.ToString(), out int result)) return result;

            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
