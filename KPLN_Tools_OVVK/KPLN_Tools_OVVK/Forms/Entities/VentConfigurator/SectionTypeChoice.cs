using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class SectionTypeChoice
    {
        public string FamilyName
        { get; private set; }

        public string TypeName
        { get; private set; }

        public string DisplayName
        {
            get
            {
                return FamilyName + " : " + TypeName;
            }
        }

        internal bool IsEmpty
        {
            get
            {
                return string.Equals(TypeName, "Пустой блок", StringComparison.OrdinalIgnoreCase);
            }
        }

        private bool IsEquipment(string name)
        {
            Func<string, string> normalize = value => Regex.Replace((value ?? string.Empty).Replace('_', ' ').Replace('ё', 'е'), @"\s+", " ").Trim();
            string family = " " + normalize(FamilyName) + " ", type = normalize(TypeName);
            var aliases = new[] { name, string.Join(" ", name.Split(' ').Reverse()) };
            return aliases.Any(alias => family.IndexOf(" Секция " + alias + " ", StringComparison.OrdinalIgnoreCase) >= 0
                || string.Equals(family.Trim(), alias, StringComparison.OrdinalIgnoreCase)
                || string.Equals(type, alias, StringComparison.OrdinalIgnoreCase)
                || type.StartsWith(alias + " ", StringComparison.OrdinalIgnoreCase));
        }

        internal ConnectionKind Connections
        {
            get
            {
                if (IsEquipment("Водяной нагреватель"))
                    return ConnectionKind.WaterHeater;
                if (IsEquipment("Водяной охладитель"))
                    return ConnectionKind.WaterCooler;
                if (IsEquipment("Фреоновый охладитель"))
                    return ConnectionKind.RefrigerantCooler;
                return ConnectionKind.None;
            }
        }

        internal string Key
        {
            get
            {
                return FamilyName + "\n" + TypeName;
            }
        }

        internal SectionTypeChoice(string familyName, string typeName)
        {
            FamilyName = familyName;
            TypeName = typeName;
        }
    }
}
