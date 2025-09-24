using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
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
            UIApplication uiapp = commandData.Application;

            CommandCheck = new CheckLinks().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

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

            ExecuteByUIApp<CheckLinks>(uiapp, false, true, true, true, true);

            return Result.Succeeded;
        }
    }
}