using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Core;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckMEPHeight : AbstrCommand, IExternalCommand
    {
        public CommandCheckMEPHeight() : base() { }

        /// <summary>
        /// Реализация IExternalCommand
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            CommandCheck = new CheckMEPHeight().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

            CheckResultStatus checkResultStatus = ExecuteByUIApp<CheckMEPHeight>(uiapp, false, true, true, true, true);
            switch (checkResultStatus)
            {
                case (CheckResultStatus.Succeeded):
                    return Result.Succeeded;
                case (CheckResultStatus.Cancelled):
                    return Result.Cancelled;
            }

            return Result.Failed;
        }
    }
}