using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Forms;
using System;
using System.Windows.Interop;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class CommandDBManager : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                DBManager mainWindow = new DBManager();
                WindowInteropHelper helper = new WindowInteropHelper(mainWindow)
                {
                    Owner = commandData.Application.MainWindowHandle
                };
                mainWindow.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                PrintError(ex);

                return Result.Failed;
            }
        }
    }
}
