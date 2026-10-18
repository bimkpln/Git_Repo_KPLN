using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class SystemNameResolver
    {
        private static readonly IComparer<string> _nameComparer = Comparer<string>.Create(CompareNames);

        internal static string Normalize(string value, string separator = "/")
        {
            if (string.IsNullOrEmpty(value))
                return value;

            return string.Join(separator, value.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries)
                .Select(name => name.Trim()).Where(name => name.Length > 0)
                .Distinct(StringComparer.Ordinal).OrderBy(name => name, _nameComparer));
        }

        // Сравниваем числовые части без преобразования в int: Т2 идёт перед Т10.
        private static int CompareNames(string left, string right)
        {
            int i = 0, j = 0;
            while (i < left.Length && j < right.Length)
            {
                if (left[i] >= '0' && left[i] <= '9' && right[j] >= '0' && right[j] <= '9')
                {
                    int endLeft = i, endRight = j;
                    while (endLeft < left.Length && left[endLeft] >= '0' && left[endLeft] <= '9') endLeft++;
                    while (endRight < right.Length && right[endRight] >= '0' && right[endRight] <= '9') endRight++;
                    while (i < endLeft && left[i] == '0') i++;
                    while (j < endRight && right[j] == '0') j++;
                    int comparison = (endLeft - i).CompareTo(endRight - j);
                    if (comparison == 0)
                        comparison = string.CompareOrdinal(left, i, right, j, endLeft - i);
                    if (comparison != 0)
                        return comparison;
                    i = endLeft;
                    j = endRight;
                }
                else
                {
                    int comparison = left[i].CompareTo(right[j]);
                    if (comparison != 0)
                        return comparison;
                    i++;
                    j++;
                }
            }
            int remaining = (left.Length - i).CompareTo(right.Length - j);
            return remaining != 0 ? remaining : string.CompareOrdinal(left, right);
        }

        internal static string Resolve(string abbreviations, IEnumerable<KeyValuePair<string, string>> systemTypes)
        {
            if (string.IsNullOrWhiteSpace(abbreviations))
                return "Нет системы";

            var requested = new HashSet<string>(abbreviations.Split(',')
                .Select(value => value.Trim()).Where(value => value.Length > 0), StringComparer.Ordinal);
            var matched = new HashSet<string>(StringComparer.Ordinal);
            var matchedNames = new Dictionary<string, string>(StringComparer.Ordinal);
            var names = new List<string>();

            foreach (KeyValuePair<string, string> systemType in systemTypes)
            {
                if (!requested.Contains(systemType.Key) || string.IsNullOrWhiteSpace(systemType.Value))
                    continue;

                if (matchedNames.TryGetValue(systemType.Key, out string previousName)
                    && !string.Equals(previousName, systemType.Value, StringComparison.Ordinal))
                    throw new InvalidOperationException("Сокращение «" + systemType.Key
                        + "» задано у разных типов систем: «" + previousName + "» и «" + systemType.Value
                        + "». Исправь сокращения типов, чтобы однозначно определить систему элемента.");

                matchedNames[systemType.Key] = systemType.Value;
                matched.Add(systemType.Key);
                if (!names.Contains(systemType.Value))
                    names.Add(systemType.Value);
            }

            string[] missing = requested.Except(matched).ToArray();
            if (missing.Length > 0 || names.Count == 0)
                throw new InvalidOperationException("Не найдены типы систем для сокращений: "
                    + (missing.Length > 0 ? string.Join(", ", missing) : abbreviations));

            return Normalize(string.Join("/", names));
        }
    }
}
