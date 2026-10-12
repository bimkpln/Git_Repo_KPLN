using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal static class ValveControl
    {
        internal const string ParameterName = "Соединитель_Приточный_Клапан";

        internal static bool? ConstantFormula(string formula)
        {
            string value = Regex.Replace(formula ?? "", @"\s+", "");
            return value == "1=1" ? (bool?)true : value == "1=0" ? (bool?)false : null;
        }

        internal static bool CanEdit(ParameterDependency parameter)
        {
            return parameter == null || !parameter.IsReadOnly && !parameter.IsReporting
            && (string.IsNullOrWhiteSpace(parameter.Formula) || ConstantFormula(parameter.Formula).HasValue);
        }
    }
}
