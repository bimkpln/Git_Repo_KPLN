using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Forms;
using System;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    internal class CommandCollectionFamilyParameters : IExternalCommand
    {
        private static CollectionFamilyParametersWindow _window;

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            try
            {
                Document document =
                    commandData.Application.ActiveUIDocument.Document;

                if (!document.IsFamilyDocument)
                {
                    TaskDialog.Show(
                        "Анализ параметров семейства",
                        "Команда доступна только в редакторе семейств.");

                    return Result.Cancelled;
                }

                if (_window != null)
                {
                    if (_window.WindowState == WindowState.Minimized)
                    {
                        _window.WindowState = WindowState.Normal;
                    }

                    _window.Activate();
                    return Result.Succeeded;
                }

                CollectionFamilyParametersWindow window =
                    new CollectionFamilyParametersWindow(commandData);

                _window = window;
                window.Closed += Window_Closed;

                WindowInteropHelper helper = new WindowInteropHelper(window);
                helper.Owner = commandData.Application.MainWindowHandle;

                window.Show();
                window.Activate();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                _window = null;
                message = ex.Message;
                return Result.Failed;
            }
        }

        private static void Window_Closed(object sender, EventArgs e)
        {
            CollectionFamilyParametersWindow window =
                sender as CollectionFamilyParametersWindow;

            if (window != null)
            {
                window.Closed -= Window_Closed;
            }

            _window = null;
        }
    }
}