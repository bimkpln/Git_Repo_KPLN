using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.Forms;
using System;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_CoordiantorAI.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class SettingsAI : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                SettingsWindow settingsWindow = new SettingsWindow();
                SetRevitOwner(settingsWindow);
                settingsWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private static void SetRevitOwner(Window window)
        {
            if (window == null || ModuleData.RevitMainWindowHandle == IntPtr.Zero)
                return;

            WindowInteropHelper helper = new WindowInteropHelper(window);
            helper.Owner = ModuleData.RevitMainWindowHandle;
        }
    }
}