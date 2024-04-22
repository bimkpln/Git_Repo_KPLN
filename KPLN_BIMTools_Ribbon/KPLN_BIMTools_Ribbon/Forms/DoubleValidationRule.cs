using System.Globalization;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public class DoubleValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (double.TryParse(value as string, NumberStyles.Any, cultureInfo, out double result))
                return ValidationResult.ValidResult;
            else
                return new ValidationResult(false, "Введите числовое значение.");
        }
    }
}
