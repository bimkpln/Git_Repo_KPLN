using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ViewsAndLists_CopyByLevels.Common
{
    /// <summary>Правила без зависимости от Revit: проверяются отдельными тестами.</summary>
    public static class NamingRules
    {
        private static readonly Regex Token = new Regex(@"\{([^{}]+)\}");

        public static string Expand(string pattern, string floor, string level, string elevation, string group, string source, int index)
        {
            var values = new Dictionary<string, string>
            {
                { "Этаж", floor }, { "Уровень", level }, { "Отметка", elevation },
                { "Типовые", group }, { "Исходное", source }, { "Индекс", index.ToString(CultureInfo.InvariantCulture) }
            };

            return Token.Replace(pattern ?? "", match =>
            {
                string value;
                if (!values.TryGetValue(match.Groups[1].Value, out value))
                    throw new ArgumentException("Неизвестная подстановка: " + match.Value);
                return value ?? "";
            }).Trim();
        }

        public static string Elevation(double meters)
        {
            return meters.ToString("+0.000;-0.000;0.000", CultureInfo.InvariantCulture);
        }

        public static bool[] SelectLevels(IList<double> elevations, double source, double gap)
        {
            if (gap <= 0 || double.IsNaN(gap) || double.IsInfinity(gap))
                throw new ArgumentException("Минимальный шаг должен быть положительным числом.");

            var result = new bool[elevations.Count];
            double previous = source;
            foreach (int i in Enumerable.Range(0, elevations.Count).OrderBy(i => elevations[i]))
            {
                if (elevations[i] > source + 1e-6 && elevations[i] - previous >= gap - 1e-6)
                {
                    result[i] = true;
                    previous = elevations[i];
                }
            }
            return result;
        }

        public static string GuessFloor(string levelName, int fallback)
        {
            Match match = Regex.Match(levelName ?? "", @"(?:этаж|эт)[_\s.\-]*(-?\d+)", RegexOptions.IgnoreCase);
            if (!match.Success)
                match = Regex.Match(levelName ?? "", @"(?:^|[_\s])(-?\d+)[_\s.-]*(?:этаж|эт)", RegexOptions.IgnoreCase);

            int value;
            return match.Success && int.TryParse(match.Groups[1].Value, out value)
                ? value.ToString(CultureInfo.InvariantCulture)
                : fallback.ToString(CultureInfo.InvariantCulture);
        }
    }
}
