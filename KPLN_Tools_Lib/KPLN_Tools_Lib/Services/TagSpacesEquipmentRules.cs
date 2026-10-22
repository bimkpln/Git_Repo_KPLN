using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Tools_Lib.Services
{
    /// <summary>Префиксы именно имени семейства. Для расширения перечня дополнять этот словарь.</summary>
    internal static class TagSpacesEquipmentRules
    {
        internal static readonly Dictionary<BuiltInCategory, string[]> FamilyPrefixes =
            new Dictionary<BuiltInCategory, string[]>
            {
                // Радиаторы — категория «Оборудование».
                { BuiltInCategory.OST_MechanicalEquipment, new[] { "554_", "556_", "557_" } },

                // Решётки — категория «Воздухораспределители».
                { BuiltInCategory.OST_DuctTerminal, new[] { "440_", "443_", "444_" } }
            };

        internal static bool MatchesFamilyName(BuiltInCategory category, string familyName)
        {
            string[] prefixes;
            return !string.IsNullOrEmpty(familyName) && FamilyPrefixes.TryGetValue(category, out prefixes)
                && prefixes.Any(prefix => familyName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }
    }
}
