using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using static KPLN_Tools_OVVK.Services.VentConfigurator.VentConfiguratorSettings;

namespace KPLN_Tools_OVVK.Services.VentConfigurator
{
    internal sealed class FamilyLoadOptions : IFamilyLoadOptions
    {
        private readonly string _rootFamilyName;
        private bool? _overwriteAccepted;

        internal bool Cancelled
        { get; private set; }

        internal FamilyLoadOptions(string rootFamilyName)
        {
            _rootFamilyName = rootFamilyName;
        }

        public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            bool accepted = ConfirmOverwrite();
            overwriteParameterValues = accepted;
            return accepted;
        }

        private bool ConfirmOverwrite()
        {
            if (_overwriteAccepted.HasValue)
                return _overwriteAccepted.Value;
            var dialog = new TaskDialog(PluginName)
            {
                MainInstruction = "В проекте уже есть это семейство. Обновить его?",
                MainContent = "Будут обновлены определение семейства и значения типов из загружаемого семейства. Это может изменить уже размещённые экземпляры.",
                CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                DefaultButton = TaskDialogResult.No
            };
            bool accepted = dialog.Show() == TaskDialogResult.Yes;
            _overwriteAccepted = accepted;
            Cancelled = !accepted;
            return accepted;
        }

        public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
            out FamilySource source, out bool overwriteParameterValues)
        {
            // Основное семейство само тоже может быть общим: для него разрешаем
            // обновление только после того же подтверждения, что и для обычного.
            if (string.Equals(sharedFamily.Name, _rootFamilyName, StringComparison.OrdinalIgnoreCase))
            {
                bool accepted = ConfirmOverwrite();
                source = accepted ? FamilySource.Family : FamilySource.Project;
                overwriteParameterValues = accepted;
                return accepted;
            }
            // Существующие общие вложенные семейства проекта сохраняем.
            source = FamilySource.Project;
            overwriteParameterValues = false;
            return true;
        }
    }
}
