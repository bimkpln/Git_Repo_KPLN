using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UIFactory;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowManager : IExternalCommand
    {
        public static GetActiveDocumentHandler ActiveDocHandler { get; private set; }
        public static ExternalEvent ExtEvent { get; private set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ActiveDocHandler = new GetActiveDocumentHandler();
            ExtEvent = ExternalEvent.Create(ActiveDocHandler);
            ExtEvent.Raise();

            try
            {
                DBProject dBProject = null;
                UIDocument uidoc = commandData.Application.ActiveUIDocument;

                if (uidoc != null)
                {
                    Document doc = uidoc.Document;
                    string fileFullName = KPLN_Looker.Module.GetFileFullName(doc);
                    dBProject = DBMainService.ProjectDbService.GetDBProject_ByRevitDocFileNameANDRVersion(fileFullName, ModuleData.RevitVersion);
                }

                if (uidoc == null || dBProject == null)
                {
                    // Для пользователей бим-отдела - показываю все проекты, включая архивные
                    bool isBIMUser = DBMainService.CurrentUserDBSubDepartment.Id == 8;

                    ElementSinglePick selectedProjectForm = SelectDbProject.CreateForm(null, ModuleData.RevitVersion, isBIMUser);
                    if ((bool)selectedProjectForm.ShowDialog())
                        dBProject = (DBProject)selectedProjectForm.SelectedElement.Element;
                    else
                        return Result.Cancelled;
                }

                ReportManagerForm mainForm = new ReportManagerForm(dBProject);
                mainForm.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                PrintError(ex);
                return Result.Cancelled;
            }
        }
    }
}
