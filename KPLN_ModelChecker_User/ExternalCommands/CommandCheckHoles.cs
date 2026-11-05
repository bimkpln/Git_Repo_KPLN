using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckHoles : AbstrCommand, IExternalCommand
    {
        public CommandCheckHoles() : base() { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            CommandCheck = new CheckHoles().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

            CheckResultStatus status = ExecuteByUIApp<CheckHoles>(uiapp);

            if (status == CheckResultStatus.Cancelled)
                return Result.Cancelled;
            return status == CheckResultStatus.Failed ? Result.Failed : Result.Succeeded;
        }
    }
}
