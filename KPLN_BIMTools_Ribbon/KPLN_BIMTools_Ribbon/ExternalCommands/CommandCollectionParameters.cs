using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Forms;
using System;
using System.Windows.Interop;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class CommandCollectionParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                CollectionParametersWindow window = new CollectionParametersWindow(commandData);
                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = commandData.Application.MainWindowHandle;
                window.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}