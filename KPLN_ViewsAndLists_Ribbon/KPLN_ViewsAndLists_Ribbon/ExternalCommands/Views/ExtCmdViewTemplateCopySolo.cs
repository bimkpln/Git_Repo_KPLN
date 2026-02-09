using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Forms;
using System;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    class CommandViewTemplateCopySolo : IExternalCommand
    {
        private static CopyViewSoloForm _window;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;

            if (_window != null && _window.IsLoaded)
            {
                BringToFront(_window);
                return Result.Succeeded;
            }

            _window = new CopyViewSoloForm(uiapp);

            try
            {
                var helper = new WindowInteropHelper(_window);
                helper.Owner = uiapp.MainWindowHandle;
            }
            catch {}


            _window.Closed += (_, __) => _window = null;

            _window.Show();
            BringToFront(_window);

            return Result.Succeeded;
        }

        private static void BringToFront(Window w)
        {
            if (w == null) return;

            if (w.WindowState == WindowState.Minimized)
                w.WindowState = WindowState.Normal;

            w.Activate();
            w.Topmost = true;
            w.Topmost = false;
            w.Focus();
        }
    }
}