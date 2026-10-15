using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Tools_OVVK.Forms.Entities;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class SystemAbbreviationValidation
    {
        // Использование стандартных типов — отдельная ошибка, сокращения у них не проверяем.
        private static readonly HashSet<string> _standardSystemTypeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Вентиляционное отверстие",
            "Водяная система пожаротушения",
            "Газовая система пожаротушения",
            "Дренчерная система пожаротушения",
            "Другие системы пожаротушения",
            "Канализация",
            "Обратная жидкость",
            "Приточная жидкость",
            "Горячее водоснабжение (внутренние сети)",
            "Холодное водоснабжение (внутренние сети)",
            "Отработанный воздух",
            "Приточный воздух",
            "Рециркулирующий воздух",
            "Удаляемый воздух",
        };

        internal static bool MatchesTypeName(string abbreviation, string typeName)
        {
            if (string.IsNullOrWhiteSpace(abbreviation) || string.IsNullOrWhiteSpace(typeName))
                return false;

            // После полного имени типа в конце сокращения обязателен "_".
            return abbreviation.EndsWith("_", StringComparison.Ordinal)
                && HasTypeSuffix(abbreviation.Substring(0, abbreviation.Length - 1), typeName);
        }

        internal static string[] GetErrors(Document doc, out string[] forbiddenTypeErrors, out OVVK_SystemManagerRowM[] issueRows)
        {
            var errors = new List<string>();
            var forbidden = new List<string>();
            var rows = new List<OVVK_SystemManagerRowM>();
            MEPSystemType[] invalidTypes;
            using (var ducts = new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystemType)))
            using (var pipes = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)))
            {
                invalidTypes = ducts.Cast<MEPSystemType>().Concat(pipes.Cast<MEPSystemType>())
                    .Where(type => _standardSystemTypeNames.Contains(type.Name) || !MatchesTypeName(type.Abbreviation, type.Name))
                    .OrderBy(type => type.Name, StringComparer.Ordinal).ToArray();
            }

            // Сети корректных типов обходить не нужно: ID элементов нужны только для строк ошибок.
            var usedElements = invalidTypes.ToDictionary(type => type.Id, type => new HashSet<ElementId>());
            if (usedElements.Count > 0)
            {
                using (var systems = new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystem)))
                {
                    foreach (MechanicalSystem system in systems)
                        if (usedElements.TryGetValue(system.GetTypeId(), out HashSet<ElementId> ids))
                            CollectElements(system, system.DuctNetwork, ids);
                }
                using (var systems = new FilteredElementCollector(doc).OfClass(typeof(PipingSystem)))
                {
                    foreach (PipingSystem system in systems)
                        if (usedElements.TryGetValue(system.GetTypeId(), out HashSet<ElementId> ids))
                            CollectElements(system, system.PipingNetwork, ids);
                }
            }

            foreach (MEPSystemType type in invalidTypes)
            {
                HashSet<ElementId> ids = usedElements[type.Id];
                if (ids.Count == 0)
                    continue;

                string category = type is PipingSystemType ? "Трубопроводная система" : "Воздуховодная система";
                if (_standardSystemTypeNames.Contains(type.Name))
                {
                    string message = $"{category} «{type.Name}»: нельзя использовать системные системы. Назначь элементам пользовательский тип системы.";
                    forbidden.Add(message);
                    rows.Add(new OVVK_SystemManagerRowM(OVVK_SystemManagerRowKind.ForbiddenSystemTypeError,
                        type.Id, type.Name, message, ids.ToArray()));
                }
                else
                {
                    string abbreviation = type.Abbreviation;
                    string value = string.IsNullOrWhiteSpace(abbreviation) ? "<пусто>" : abbreviation;
                    string message = $"{category} «{type.Name}»: «Сокращение» = «{value}» — должно завершаться полным именем типа и «_».";
                    errors.Add(message);
                }
            }
            forbiddenTypeErrors = forbidden.ToArray();
            issueRows = rows.ToArray();
            return errors.ToArray();
        }
        private static void CollectElements(MEPSystem system, ElementSet network, HashSet<ElementId> ids)
        {
            foreach (Element element in system.Elements)
                ids.Add(element.Id);
            if (network != null)
                foreach (Element element in network)
                    ids.Add(element.Id);
            if (system.BaseEquipment != null)
                ids.Add(system.BaseEquipment.Id);
        }

        private static bool HasTypeSuffix(string value, string typeName)
        {
            return string.Equals(value, typeName, StringComparison.Ordinal)
                || value.EndsWith("_" + typeName, StringComparison.Ordinal);
        }
    }
}
