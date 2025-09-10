using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.WPFItems;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckLinks : AbstrCommand, IExternalCommand
    {
        public CommandCheckLinks() : base() { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            CommandCheck = new CheckLinks();
            ElemsToCheck = CommandCheck.GetElemsToCheck(commandData.Application.ActiveUIDocument.Document);

            // Блокирую проверку части линков ПРИ РУЧНОМ ЗАПУСКЕ
            // Для авт. запуска блок не нужен, достаточно проверять часть.
            // Даже если это перезапуск - изначально всё равно нужно вручную запустить и подгрузить линки
            foreach (RevitLinkInstance rli in ElemsToCheck.Cast<RevitLinkInstance>())
            {
                Document linkDoc = rli.GetLinkDocument();
                if (linkDoc == null)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = $"Необходимо загрузить ВСЕ связи. Проверь диспетчер Revit-связей",
                    };
                    taskDialog.Show();

                    return Result.Cancelled;
                }
            }

            if (ExecuteByUIApp<CheckLinks>(commandData.Application, true, true, true, true))
                return Result.Succeeded;

            return Result.Cancelled;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}