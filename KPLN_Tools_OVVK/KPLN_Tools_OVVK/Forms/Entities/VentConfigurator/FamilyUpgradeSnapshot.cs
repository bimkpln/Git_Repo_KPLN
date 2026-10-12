using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyUpgradeSnapshot
    {
        internal readonly Dictionary<string, FamilyUpgradeParameter> Parameters = new Dictionary<string, FamilyUpgradeParameter>(StringComparer.Ordinal);
        internal readonly Dictionary<string, Dictionary<string, object>> Types = new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);

        internal static void Verify(FamilyUpgradeSnapshot before, FamilyUpgradeSnapshot source, FamilyUpgradeSnapshot after)
        {
            var errors = new List<string>();
            foreach (var parameter in source.Parameters)
            {
                FamilyUpgradeParameter current;
                if (!after.Parameters.TryGetValue(parameter.Key, out current)
                    || !parameter.Value.SameDefinition(current, parameter.Key == ValveControl.ParameterName))
                    errors.Add("Не обновлено определение параметра: " + parameter.Key);
            }
            foreach (string name in source.Types.Keys)
                if (!after.Types.ContainsKey(name))
                    errors.Add("Не добавлен тип новой основы: " + name);
            foreach (var type in before.Types)
            {
                Dictionary<string, object> values;
                if (!after.Types.TryGetValue(type.Key, out values))
                {
                    errors.Add("Потерян тип: " + type.Key);
                    continue;
                }
                foreach (var input in type.Value)
                {
                    FamilyUpgradeParameter parameter;
                    // Удалённые параметры и новые формулы берутся из новой основы.
                    // Совместимые редактируемые значения должны сохраниться.
                    if (!after.Parameters.TryGetValue(input.Key, out parameter) || !parameter.Writable
                        || before.Parameters[input.Key].Kind != parameter.Kind)
                        continue;
                    object value;
                    bool same = values.TryGetValue(input.Key, out value) && (value is double && input.Value is double
                        ? Math.Abs((double)value - (double)input.Value) < 1e-9 : Equals(value, input.Value));
                    if (!same)
                        errors.Add(type.Key + " / " + input.Key);
                }
            }
            if (errors.Count != 0)
                throw new InvalidOperationException("Проверка обновления семейства не пройдена. Файл проекта не заменён.\n"
                + string.Join("\n", errors.Take(25)));
        }

        internal string ChangesFrom(FamilyUpgradeSnapshot before)
        {
            var added = Parameters.Keys.Except(before.Parameters.Keys).ToList();
            var removed = before.Parameters.Keys.Except(Parameters.Keys).ToList();
            var changed = Parameters.Keys.Intersect(before.Parameters.Keys)
                .Where(k => !Parameters[k].SameDefinition(before.Parameters[k], false)).ToList();
            Func<IEnumerable<string>, string> list = names => string.Join(", ", names.Take(12));
            return "Новых параметров: " + added.Count + (added.Count > 0 ? " (" + list(added) + ")" : "")
                + "\nИзменённых определений и формул: " + changed.Count + (changed.Count > 0 ? " (" + list(changed) + ")" : "")
                + "\nУдалённых из основы параметров: " + removed.Count + (removed.Count > 0 ? " (" + list(removed) + ")" : "");
        }
    }
}
