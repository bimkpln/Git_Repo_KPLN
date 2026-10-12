using System;

namespace KPLN_Tools_OVVK.Forms.Entities.VentConfigurator
{
    internal sealed class FamilyActionPlan
    {
        internal string Path;
        internal bool WritesFamily, NeedsPath;

        internal static FamilyActionPlan Make(RequestKind kind, string savedPath, string configuredPath, string persistedName, bool dirty)
        {
            bool write = kind == RequestKind.Save || kind == RequestKind.DeleteType;
            if (!write && kind != RequestKind.OpenInRevit && kind != RequestKind.AddToProject)
                throw new ArgumentException("Неизвестное действие.");
            if (!write && (dirty || string.IsNullOrWhiteSpace(savedPath) || string.IsNullOrWhiteSpace(persistedName)))
                throw new InvalidOperationException("Сначала сохраните тип и изменения кнопкой «Сохранить».");
            // Для определённого проекта место записи всегда задаёт его папка.
            // Открытие уже сохранённого типа использует его фактический файл.
            string path = write && !string.IsNullOrWhiteSpace(configuredPath) ? configuredPath : savedPath;
            return new FamilyActionPlan { WritesFamily = write, Path = path, NeedsPath = write && string.IsNullOrWhiteSpace(path) };
        }
    }
}
