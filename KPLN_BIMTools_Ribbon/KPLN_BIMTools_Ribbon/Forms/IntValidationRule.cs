using System.Globalization;
using System.Windows.Controls;

namespace KPLN_BIMTools_Ribbon.Forms
{
    public class IntValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (int.TryParse(value as string, NumberStyles.Any, cultureInfo, out int result))
                return ValidationResult.ValidResult;
            else
                return new ValidationResult(false, "Введите числовое значение.");
        }
    }
}
