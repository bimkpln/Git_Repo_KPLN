using Autodesk.Revit.UI;
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
            anyErrors |= Run<CommandCheckLinks>(uiApp);
            anyErrors |= Run<CommandCheckWorksets>(uiApp);
            // anyErrors |= Run<CommandCheckSomethingElse>(uiApp);

            if (anyErrors)
            {
                MessageBox.Show(
                    "Вы произвели синхронизацию проекта, в котором находятся критические ошибки.\n\n" +
                    $"Окна с ошибками появились отдельно, чтобы они больше не всплывали - исправь замечания",
                    "KPLN: Ошибка",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Запуск выбранной проверки
        /// </summary>
        private static bool Run<T>(UIApplication uiApp) where T : AbstrCommand<T>, new()
        {
            var cmd = new T();

            // Ніякіх Activity-логів, ніякіх "усё добра" банэраў — ціхі запуск
            cmd.ExecuteByUIApp(
                uiapp: uiApp,
                setPluginActivity: false,
                setLastRun: false,
                showMainForm: false,
                showSuccsessText: false);

            // ДОСТУП да вынікаў праз generic-static (асобнае статычнае поле на кожны closed generic T)
            var entities = AbstrCommand<T>.CheckerEntities;
            bool hasErrors = entities != null && entities.Any(e => e.Status == KPLN_ModelChecker_Lib.ErrorStatus.Error);

            // Паказваем справаздачу толькі калі ёсць памылкі
            if (hasErrors)
                cmd.ReportCreatorAndDemonstrator(uiApp);

            return hasErrors;
        }
    }
}
