using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Forms;
using KPLN_ViewsAndLists_Ribbon.Forms.ViewModels;
using System;
using System.Collections.Generic;
using System.Windows.Interop;

namespace KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class ExtCmdScheduleManager : IExternalCommand
    {
        public const string PluginName = "Конструктор спецификаций";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null || uidoc.Document.IsFamilyDocument)
            {
                TaskDialog.Show(PluginName, "Откройте проект Revit.");
                return Result.Cancelled;
            }

            try
            {
                var selected = new HashSet<ElementId>(uidoc.Selection.GetElementIds());
                if (selected.Count == 0 && uidoc.ActiveView is ViewSchedule) selected.Add(uidoc.ActiveView.Id);
                var window = new ScheduleManagerForm(new ScheduleManagerVM(uidoc.Document, selected));
                new WindowInteropHelper(window).Owner = commandData.Application.MainWindowHandle;
                window.ShowDialog();
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
