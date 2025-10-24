using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;
using KPLN_ModelChecker_User.ExternalCommands;
using System.Linq;
using System.Windows.Forms;

namespace KPLN_Looker.Services
{
    public static class CheckBatchRunner
    {
        /// <summary>
        /// Запуск всех проверок по преднастройке в коде
        /// </summary>
        public static void RunAll(UIApplication uiapp, string docTitle)
        {
            bool anyErrors = false;

            // Общий фильтр по проектам
            if (!docTitle.StartsWith("ИЗМЛ_")
                && !docTitle.StartsWith("ИЗМЛ23_")
                && !docTitle.StartsWith("ПШМ1_")
                && !docTitle.StartsWith("ПСРВ_")
                && !docTitle.StartsWith("SH1-"))
            {
                // anyErrors |= Run<CommandCheckSomethingElse1, CommandCheckSomethingElse2>(uiapp);

                anyErrors |= Run<CommandCheckLinks, CheckLinks>(uiapp);
                anyErrors |= Run<CommandCheckWorksets, CheckWorksets>(uiapp);

                // Персональный фильтр по проектам
                if (!docTitle.StartsWith("МТРС_")
                    && !docTitle.StartsWith("СЕТ_1_"))
                    anyErrors |= Run<CommandCheckMainLines, CheckMainLines>(uiapp);
            }

            if (anyErrors)
            {
                MessageBox.Show(
                    "Вы произвели синхронизацию проекта, в котором находятся критические ошибки.\n\n" +
                    $"Окна с ошибками появились отдельно, чтобы они больше не всплывали - исправь замечания (как минимум категории \"Ошибка\", т.е. красного цвета)",
                    "KPLN: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Запуск выбранной проверки
        /// </summary>
        private static bool Run<TCmd, TCheck>(UIApplication uiapp)
            where TCmd : AbstrCommand, new()
            where TCheck : AbstrCheck, new()
        {
            var cmd = new TCmd();

            cmd.ExecuteByUIApp<TCheck>(
                uiapp: uiapp,
                onlyErrorType: true,
                setPluginActivity: false,
                setLastRun: false,
                showMainForm: false,
                showSuccsessText: false);

            // ДОСТУП да вынікаў праз generic-static (асобнае статычнае поле на кожны closed generic T)
            // ВАЖНО: Если будет ПАКЕТНАЯ проверка между файлами - нужны копии результатов таких проверок,
            // т.к. инстанс AbstrCommand один, результаты просто перетираются при запуске на новом файле
            var entities = cmd.CommandCheck.CheckerEntitiesColl;
            bool hasErrors = entities != null && entities.Any();

            // Паказваем справаздачу толькі калі ёсць памылкі
            if (hasErrors)
                cmd.ReportCreatorAndDemonstrator<TCheck>(uiapp);

            return hasErrors;
        }
    }
}
