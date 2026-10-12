using System;
using System.Linq;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal static class VentTypeName
    {
        internal static string ValidateTypeName(string value)
        {
            string name = (value ?? string.Empty).Trim();
            if (name.Length == 0)
                throw new ArgumentException("Введите имя нового типа.");
            if (string.Equals(name, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Это имя базового типа. Введите другое имя для новой установки.");
            if (name.Any(char.IsControl) || name.IndexOfAny(new[] { '\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' }) >= 0)
                throw new ArgumentException("Имя типа содержит недопустимые символы: \\ : { } [ ] | ; < > ? ` ~ или перевод строки.");
            return name;
        }
    }
}
