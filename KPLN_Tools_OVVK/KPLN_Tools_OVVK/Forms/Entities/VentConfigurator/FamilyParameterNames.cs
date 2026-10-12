using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    // Нумерация физического последнего блока остаётся 12. В другой редакции RFA его параметры могут иметь номер 2.
    // «Секция_2_Соединитель_Клапан_…» всегда обозначает отдельный второй блок, не последний соединитель.
    internal static class FamilyParameterNames
    {
        internal static string Canonical(string name)
        {
            string result = Regex.Replace((name ?? "").Replace("Соеденитель", "Соединитель"), @"Смещение_по_([XY])", "Смещение по $1");
            result = Regex.Replace(result, @"^(Вода|Фреон|Дренаж)(_Патрубок_.*Смещение)_([XY])$", "$1$2 $3");
            if (result == "Дренаж_Патрубок_Смещение")
                result += " Y";
            return Regex.Replace(result, @"Секция_2_Соединитель_(?!Клапан_)", "Секция_12_Соединитель_");
        }

        internal static IEnumerable<string> Aliases(string name)
        {
            string canonical = Canonical(name);
            foreach (string numbered in new[] { canonical, canonical.Replace("Секция_12_Соединитель_", "Секция_2_Соединитель_") }.Distinct())
                foreach (string spelling in new[] { numbered, numbered.Replace("Соединитель", "Соеденитель") }.Distinct())
                    foreach (string spacing in new[] { spelling, spelling.Replace("Смещение по ", "Смещение_по_") }.Distinct())
                        yield return spacing;
            if (Regex.IsMatch(canonical, @"^(Вода|Фреон|Дренаж)_Патрубок_.*Смещение [XY]$"))
                yield return Regex.Replace(canonical, @"Смещение ([XY])$", "Смещение_$1");
            if (canonical == "Дренаж_Патрубок_Смещение Y")
                yield return "Дренаж_Патрубок_Смещение";
        }

        internal static bool IsOffset(string name)
        {
            return Regex.IsMatch(Canonical(name), @"^Секция_(1|12)_Соединитель_Смещение по [XY]$")
            || Regex.IsMatch(Canonical(name), @"^(Вода|Фреон|Дренаж)_Патрубок_(?:(?:Приток|Обратка)_)?Смещение [XY]$");
        }
    }
}
