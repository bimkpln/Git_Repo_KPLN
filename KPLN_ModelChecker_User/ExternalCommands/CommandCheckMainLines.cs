using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.WPFItems;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckMainLines : AbstrCommand, IExternalCommand
    {
        public CommandCheckMainLines() : base() { }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            CommandCheck = new CheckMainLines();
            ElemsToCheck = CommandCheck.GetElemsToCheck(commandData.Application.ActiveUIDocument.Document);

            ExecuteByUIApp<CheckMainLines>(commandData.Application, false, true, true, true, true);
            
            return Result.Succeeded;
        }

        private protected override void SetWPFEntityFiltration(WPFReportCreator report)
        {
            report.SetWPFEntityFiltration_ByErrorHeader();
        }
    }
}
