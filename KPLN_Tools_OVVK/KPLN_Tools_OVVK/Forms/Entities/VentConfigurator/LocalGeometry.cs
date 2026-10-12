using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    // Правила из отчёта исходного RFA. Здесь нет Revit API, открытия документов и геометрического решателя.
    // Накопленные длины — результат сложения только видимых секций; длина последнего соединителя является входом.
    internal sealed class LocalGeometry
    {
        internal readonly Dictionary<int, double?> EndsMm = new Dictionary<int, double?>();
        internal string Limitation;

        internal double? TotalLengthMm
        {
            get
            {
                double? value;
                return EndsMm.TryGetValue(12, out value) ? value : null;
            }
        }

        internal static string Format(double? value)
        {
            return value.HasValue ? value.Value.ToString("0.######", CultureInfo.CurrentCulture) : "—";
        }

        internal string Report()
        {
            return "Локальный расчёт по формулам исходного семейства:\n"
                + string.Join("\n", EndsMm.OrderBy(p => p.Key).Select(p => "Системный_" + p.Key + "_Блок_Длина = " + Format(p.Value) + " мм"))
                + "\nУстановка_Длина = " + Format(TotalLengthMm) + " мм"
                + (string.IsNullOrWhiteSpace(Limitation) ? "" : "\n" + Limitation);
        }

        internal static bool SourceFlag(SectionCatalog catalog, string name, bool fallback)
        {
            ParameterDependency p;
            if (!catalog.Dependencies.Parameters.TryGetValue(name, out p))
                return fallback;
            string formula = Compact(p.Formula);
            if (formula == "1=1")
                return true;
            if (formula == "1=0")
                return false;
            return p.IntegerValue.HasValue ? p.IntegerValue.Value != 0 : fallback;
        }

        private static string Compact(string formula)
        {
            return Regex.Replace(FamilyParameterNames.Canonical(formula), @"\s+", "");
        }

        internal static LocalGeometry Calculate(InstallationConfiguration configuration)
        {
            var result = new LocalGeometry();
            var dependencies = configuration.Catalog.Dependencies;
            var issues = new List<string>();
            Action<string, string> check = (name, expected) =>
            {
                ParameterDependency p;
                if (dependencies.Parameters.TryGetValue(name, out p) && Compact(p.Formula) != Compact(expected))
                    issues.Add("Формула «" + name + "» отличается от перенесённой в плагин: " + (p.Formula ?? "без формулы") + ".");
            };
            check("Системный_1_Блок_Длина", "if(Соединитель_Приточный, Секция_1_Соединитель_Длина, 0 мм)");
            check("Системный_2_Блок_Длина", "if(Соединитель_Приточный_Клапан, Секция_2_Соединитель_Клапан_Длина + Системный_1_Блок_Длина, Системный_1_Блок_Длина)");
            for (int slot = 3; slot <= 11; slot++)
            {
                int section = slot - 2;
                check("Системный_" + slot + "_Блок_Длина", "if(Системный_Видимость_" + section + "_Блок, "
                    + InstallationConfiguration.DimensionParameterName(slot, "Длина") + " + Системный_" + (slot - 1)
                    + "_Блок_Длина, Системный_" + (slot - 1) + "_Блок_Длина)");
                if (section > 1)
                    check("Системный_Видимость_" + section + "_Блок",
                    "if(Секции_Промежуточные_Количество > " + (section - 1) + ", 1 = 1, 1 = 0)");
            }
            check("Системный_12_Блок_Длина", "if(Соединитель_Отработанный, Секция_12_Соединитель_Длина + Системный_11_Блок_Длина, Системный_11_Блок_Длина)");
            check("Установка_Длина", "Системный_12_Блок_Длина");
            foreach (string name in new[] { "Соединитель_Приточный", "Соединитель_Отработанный", "Системный_Видимость_1_Блок", "Соединитель_Приточный_Клапан" })
            {
                ParameterDependency p;
                if (dependencies.Parameters.TryGetValue(name, out p) && !string.IsNullOrWhiteSpace(p.Formula)
                    && Compact(p.Formula) != "1=1" && Compact(p.Formula) != "1=0")
                    issues.Add("Не перенесена формула «" + name + "»: " + p.Formula + ".");
            }
            var blocks = Enumerable.Range(0, configuration.Blocks.Count).ToDictionary(configuration.SlotAt, i => configuration.Blocks[i]);
            double? end = issues.Count == 0 ? (double?)0 : null;
            for (int slot = 1; slot <= 12; slot++)
            {
                bool visible = slot == 1 ? SourceFlag(configuration.Catalog, "Соединитель_Приточный", true)
                    : slot == 12 ? SourceFlag(configuration.Catalog, "Соединитель_Отработанный", true)
                    : slot == 2 ? configuration.HasValve : slot == 3 ? SourceFlag(configuration.Catalog, "Системный_Видимость_1_Блок", true)
                    : slot - 2 <= configuration.IntermediateCount;
                if (visible)
                {
                    SectionValue block;
                    double length;
                    if (!blocks.TryGetValue(slot, out block) || !block.Length.TryNumber(out length) || length < 0)
                        end = null;
                    else if (end.HasValue)
                        end += length;
                }
                result.EndsMm[slot] = end;
            }
            var unknown = configuration.Dimensions().Where(p => p.Value.IsCalculated).Select(p => p.Key).ToList();
            if (unknown.Count > 0)
                issues.Add("Для этих размеров локальные правила не подтверждены: " + string.Join(", ", unknown) + ".");
            result.Limitation = string.Join("\n", issues);
            return result;
        }
    }
}
