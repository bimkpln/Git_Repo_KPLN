using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
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
                // Для пользователей бим-отдела - показываю все проекты, включая архивные
                bool isBIMUser = ClashesMainCollection.CurrentDBUser.SubDepartmentId == 8;
                
                ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm(isBIMUser);
                bool? dialogResult = selectedProjectForm.ShowDialog();
                if (selectedProjectForm.Status == UIStatus.RunStatus.Run)
                {
                    DBProject dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
                    ReportManagerForm mainForm = new ReportManagerForm(dBProject);
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
