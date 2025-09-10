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
        public static void RunAll(UIApplication uiApp)
        {
            bool anyErrors = false;

            // Дадавай новыя праверкі тут:
            anyErrors |= Run<CommandCheckLinks, CheckLinks>(uiApp);
            anyErrors |= Run<CommandCheckMainLines, CheckMainLines>(uiApp);
            anyErrors |= Run<CommandCheckWorksets, CheckWorksets>(uiApp);
            // anyErrors |= Run<CommandCheckSomethingElse1, CommandCheckSomethingElse2>(uiApp);

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
        private static bool Run<TCmd, TCheck>(UIApplication uiApp)
            where TCmd : AbstrCommand, new()
            where TCheck : AbstrCheck, new()
        {
            var cmd = new TCmd();

            cmd.ExecuteByUIApp<TCheck>(
                uiapp: uiApp,
                onlyErrorType: true,
                setPluginActivity: false,
                setLastRun: false,
                showMainForm: false,
                showSuccsessText: false);

            // ДОСТУП да вынікаў праз generic-static (асобнае статычнае поле на кожны closed generic T)
            var entities = cmd.CheckerEntities;
            bool hasErrors = entities != null && entities.Any();

            // Паказваем справаздачу толькі калі ёсць памылкі
            if (hasErrors)
                cmd.ReportCreatorAndDemonstrator<TCheck>(uiApp);

            return hasErrors;
        }
    }
}
