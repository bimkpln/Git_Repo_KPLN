using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KPLN_Tools_Lib.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class TagRoomsExtCmd : IExternalCommand
    {
        public const string PluginName = "Маркировать помещения";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            => TagSpacesExtCmd.Execute(commandData, ref message, elements, true);
    }
}
