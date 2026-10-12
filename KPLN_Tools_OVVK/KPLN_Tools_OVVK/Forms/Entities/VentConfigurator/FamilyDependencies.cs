using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyDependencies
    {
        internal Dictionary<string, ParameterDependency> Parameters = new Dictionary<string, ParameterDependency>(StringComparer.Ordinal);
        internal List<string> ReadWarnings = new List<string>();

        internal ParameterDependency Find(string name)
        {
            var matches = FamilyParameterNames.Aliases(name).Where(Parameters.ContainsKey).Select(n => Parameters[n]).ToList();
            if (matches.Count > 1)
                throw new InvalidOperationException("Неоднозначное соответствие параметра «" + name + "»: " + string.Join(", ", matches.Select(p => p.Name)));
            return matches.FirstOrDefault();
        }

        internal bool IsCalculated(string name)
        {
            var parameter = Find(name);
            return parameter != null
                && (parameter.IsReadOnly || parameter.IsReporting || !string.IsNullOrWhiteSpace(parameter.Formula));
        }

        internal static IEnumerable<string> FormulaReferences(string formula, IEnumerable<string> names)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return Enumerable.Empty<string>();
            // Имена ищем целиком, сначала самые длинные; строковые константы не являются ссылками.
            string unquoted = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", " ");
            string alternatives = string.Join("|", names.Where(n => !string.IsNullOrWhiteSpace(n)).Distinct()
                .OrderByDescending(n => n.Length).Select(Regex.Escape));
            if (alternatives.Length == 0)
                return Enumerable.Empty<string>();
            return Regex.Matches(unquoted, @"(?<![\p{L}\p{N}_])(?:" + alternatives + @")(?![\p{L}\p{N}_])")
                .Cast<Match>().Select(m => m.Value).Distinct().ToList();
        }

        internal void ConnectFormulas()
        {
            foreach (var parameter in Parameters.Values)
            {
                parameter.Inputs.Clear();
                parameter.Dependents.Clear();
            }
            foreach (var parameter in Parameters.Values)
                foreach (string input in FormulaReferences(parameter.Formula, Parameters.Keys))
                {
                    parameter.Inputs.Add(input);
                    Parameters[input].Dependents.Add(parameter.Name);
                }
        }

        internal string Tooltip(string name)
        {
            var parameter = Find(name);
            if (parameter == null)
                return name;
            var lines = new List<string> { parameter.Name };
            lines.Add(string.IsNullOrWhiteSpace(parameter.Formula) ? "Формула отсутствует." : "Формула: " + parameter.Formula);
            lines.Add(parameter.IsReporting ? "Отчётный размер: зависит от геометрии." : parameter.IsReadOnly ? "Только чтение." : "");
            if (parameter.Inputs.Count > 0)
                lines.Add("Зависит от: " + string.Join(", ", parameter.Inputs));
            if (parameter.Dependents.Count > 0)
                lines.Add("Влияет на: " + string.Join(", ", parameter.Dependents.Take(5))
                + (parameter.Dependents.Count > 5 ? "…" : ""));
            lines.AddRange(parameter.Links.Take(4));
            return string.Join("\n", lines);
        }

        private HashSet<string> Reachable(IEnumerable<string> seeds, bool upstream)
        {
            var visited = new HashSet<string>(StringComparer.Ordinal);
            var pending = new Stack<string>(seeds);
            while (pending.Count > 0)
            {
                string name = pending.Pop();
                ParameterDependency parameter;
                if (!visited.Add(name) || !Parameters.TryGetValue(name, out parameter))
                    continue;
                foreach (string next in upstream ? parameter.Inputs : parameter.Dependents)
                    pending.Push(next);
            }
            return visited;
        }

        internal string Report(IEnumerable<string> names)
        {
            var seeds = names.SelectMany(FamilyParameterNames.Aliases).Where(Parameters.ContainsKey).Distinct().ToList();
            var relevant = Reachable(seeds, false);
            // Включаем и остальные входы зависимых формул: например, флаги видимости и включения соединителей.
            relevant.UnionWith(Reachable(relevant.ToList(), true));
            var lines = new List<string> { "Зависимости параметров конфигуратора (снимок исходного семейства):" };
            foreach (string name in relevant.OrderBy(n => n, StringComparer.Ordinal))
            {
                ParameterDependency parameter;
                if (!Parameters.TryGetValue(name, out parameter))
                    continue;
                lines.Add(name + (string.IsNullOrWhiteSpace(parameter.Formula) ? " [без формулы]" : " = " + parameter.Formula)
                    + " [ReadOnly=" + parameter.IsReadOnly + "; Reporting=" + parameter.IsReporting
                    + "; Instance=" + parameter.IsInstance + "; исходное значение=" + parameter.SourceValue + "]");
                lines.AddRange(parameter.Links.Distinct().Select(link => "  " + link));
            }
            lines.AddRange(ReadWarnings);
            return string.Join("\n", lines);
        }
    }
}
