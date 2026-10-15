using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class SystemNameResolver
    {
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

            return string.Join("/", names);
        }
    }
}
