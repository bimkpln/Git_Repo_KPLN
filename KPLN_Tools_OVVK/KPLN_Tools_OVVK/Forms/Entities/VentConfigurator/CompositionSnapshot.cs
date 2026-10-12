using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class CompositionSnapshot
    {
        internal readonly Dictionary<int, HashSet<string>> Choices = new Dictionary<int, HashSet<string>>();
        internal readonly Dictionary<string, object> Settings = new Dictionary<string, object>(StringComparer.Ordinal);

        internal static void RequireChoices(CompositionSnapshot required, CompositionSnapshot actual, string message)
        {
            var missing = new List<string>();
            foreach (var slot in required.Choices)
            {
                HashSet<string> available;
                if (!actual.Choices.TryGetValue(slot.Key, out available))
                    available = new HashSet<string>();
                missing.AddRange(slot.Value.Where(key => !available.Contains(key))
                    .Select(key => InstallationConfiguration.ParameterName(slot.Key) + ": " + key.Replace("\n", " : ")));
            }
            if (missing.Count > 0)
                throw new InvalidOperationException(message + "\n" + string.Join("\n", missing.Take(20)));
        }

        internal static void RequireSettings(CompositionSnapshot before, CompositionSnapshot after)
        {
            var changed = new List<string>();
            foreach (var pair in before.Settings)
            {
                object current;
                bool found = after.Settings.TryGetValue(pair.Key, out current);
                bool equal = found && (pair.Value is double && current is double
                    ? Math.Abs((double)pair.Value - (double)current) < 1e-9 : Equals(pair.Value, current));
                if (!equal)
                    changed.Add(pair.Key.Replace("\n", " / "));
            }
            changed.AddRange(after.Settings.Keys.Where(key => !before.Settings.ContainsKey(key)).Select(key => key.Replace("\n", " / ")));
            if (changed.Count > 0)
                throw new InvalidOperationException("Обновление вложенных семейств меняет типы или настройки установок. Файл проекта не изменён.\n"
                    + string.Join("\n", changed.Take(20)));
        }
    }
}
