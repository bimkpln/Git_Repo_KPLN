using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckFlatsAreaCompare : AbstrCommand, IExternalCommand
    {
        public CommandCheckFlatsAreaCompare() : base() { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            CommandCheck = new CheckFlatsAreaCompare().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

            ExecuteByUIApp<CheckFlatsAreaCompare>(uiapp, false, true, true, true, true);

            return Result.Succeeded;
        }
    }
}
