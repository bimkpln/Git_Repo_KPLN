using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Library_DataBase.Collections;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowManager : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                ElementPick selectedProjectForm = SelectDbProject.CreateForm();
                bool? dialogResult = selectedProjectForm.ShowDialog();
                if (selectedProjectForm.Status == UIStatus.RunStatus.Run)
                {
                    ReportManager mainForm = new ReportManager((DbProject)selectedProjectForm.SelectedElement.Element);
                    mainForm.Show();
                }
                else
                {
                    return Result.Cancelled;
                }

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
