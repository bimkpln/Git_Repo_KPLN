using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    // Создание и редактирование различаются намерением, а не совпадением введённого имени.
    internal sealed class TypeWritePlan
    {
        internal string SourceTypeName, TargetName;
        internal bool Create;

        internal static TypeWritePlan Make(IEnumerable<string> existingNames, string targetName, string persistedName, bool freshLibraryCopy, string copySourceTypeName = null)
        {
            var names = existingNames.ToList();
            var result = new TypeWritePlan { TargetName = targetName, Create = string.IsNullOrWhiteSpace(persistedName) };
            if (string.Equals(targetName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Шаблонный тип нельзя перезаписывать. Задайте имя установки.");
            if (string.Equals(persistedName, BaseFamilyTypeName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Шаблонный тип служит только основой для новой установки.");
            if (!result.Create)
            {
                result.SourceTypeName = names.SingleOrDefault(n => string.Equals(n, persistedName, StringComparison.Ordinal));
                if (result.SourceTypeName == null)
                    throw new InvalidOperationException("Редактируемый тип «" + persistedName + "» не найден. Заново откройте семейство.");
            }
            else
            {
                string sourceName = copySourceTypeName ?? BaseFamilyTypeName;
                result.SourceTypeName = names.SingleOrDefault(n => string.Equals(n, sourceName, StringComparison.Ordinal));
                if (result.SourceTypeName == null)
                    throw new InvalidOperationException("В семействе нет исходного типа «" + sourceName + "» для создания копии. Он мог быть удалён или переименован.");
            }
            if (names.Any(n => (result.Create || n != result.SourceTypeName)
                && string.Equals(n, targetName, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Тип «" + targetName + "» уже существует. Выберите его в списке для редактирования или задайте другое имя новому типу.");
            return result;
        }
    }
}
