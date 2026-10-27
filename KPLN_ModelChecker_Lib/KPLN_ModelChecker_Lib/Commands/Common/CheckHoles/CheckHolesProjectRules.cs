using System;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    /// <summary>Уточнения по проекту для отбора отверстий и определения направления прохода.</summary>
    internal static class CheckHolesProjectRules
    {
        internal static bool IsSetProject(string documentTitle)
        {
            // Как в Looker: проект определяется при запуске по названию текущей модели.
            return documentTitle != null && documentTitle.IndexOf("СЕТ_1", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        internal static bool IsSetWallOpening(string documentTitle, string familyName, bool isWindow)
        {
            // СЕТ_1: только отверстия в стенах — семейства «Отверстие» категории «Окна».
            // Шахты в оборудовании не входят в область этой проверки.
            return IsSetProject(documentTitle) && isWindow && !string.IsNullOrEmpty(familyName)
                && familyName.IndexOf("Отверстие", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
