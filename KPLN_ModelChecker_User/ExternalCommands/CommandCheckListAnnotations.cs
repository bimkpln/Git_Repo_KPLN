using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using System.Linq;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckListAnnotations : AbstrCommand, IExternalCommand
    {
        public CommandCheckListAnnotations() : base() { }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            CommandCheck = new CheckListAnnotations().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

            if (ElemsToCheck.Count() > 0)
                ExecuteByUIApp<CheckListAnnotations>(uiapp, false, true, true, true, true);

            return Result.Succeeded;
        }
    }
}
