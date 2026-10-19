using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_ViewsAndLists_Ribbon.Common.Schedules
{
    /// <summary>Правила без зависимости от Revit: проверяются отдельными тестами.</summary>
    public static class ScheduleNaming
    {
        private static readonly Regex Token = new Regex(@"\{([^{}]+)\}");

        public static string Expand(string pattern, string floor, string level, string elevation, string group, string source, int index)
        {
            return Expand(pattern, floor, level, elevation, group, source, index, null);
        }

        public static string Expand(string pattern, string floor, string level, string elevation, string group, string source, int index, string filters)
        {
            return Expand(pattern, floor, level, elevation, group, source, index, filters, null, null);
        }

        public static string Expand(string pattern, string floor, string level, string elevation, string group, string source, int index, string filters, string parameter, string copyValue)
        {
            var values = new Dictionary<string, string>
            {
                { "Этаж", floor }, { "Уровень", level }, { "Отметка", elevation },
                { "Типовые", group }, { "Исходное", source }, { "Индекс", index.ToString(CultureInfo.InvariantCulture) }
            };

            if (filters != null) values.Add("Фильтры", filters);
            if (parameter != null) values.Add("Параметр", parameter);
            if (copyValue != null) values.Add("Значение", copyValue);

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
