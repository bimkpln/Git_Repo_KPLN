extern alias revit;

using KPLN_ModelChecker_Coordinator.Common;
using KPLN_ModelChecker_Coordinator.DB;
using KPLN_ModelChecker_Coordinator.Forms;
using revit.Autodesk.Revit.Attributes;
using revit.Autodesk.Revit.DB;
using revit.Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_ModelChecker_Coordinator.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowStatistics : IExternalCommand
    {
        private bool IsValid(KPLN_Library_DataBase.Collections.DbDocument doc)
        {
            if (doc == null) return false;
            if (doc.Department == null || doc.Name == null || doc.Project == null || doc.Path == null || doc.Code == null || doc.Department == null) return false;
            return true;
        }
        
        private static bool InList(KPLN_Library_DataBase.Collections.DbProject project, List<KPLN_Library_DataBase.Collections.DbProject> projects)
        {
            try
            {
                foreach (var i in projects)
                {
                    if (project.Id == i.Id)
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                List<KPLN_Library_DataBase.Collections.DbProject> projects = new List<KPLN_Library_DataBase.Collections.DbProject>();
                foreach (KPLN_Library_DataBase.Collections.DbDocument doc in KPLN_Library_DataBase.DbControll.Documents)
                {
                    if (!IsValid(doc)) continue;
                    if (File.Exists(string.Format(@"{0}\doc_id_{1}.sqlite", MyPathes.ModelCheckerDBPath, doc.Id.ToString())))
                    {
                        if (!InList(doc.Project, projects) || projects.Count == 0)
                        {
                            projects.Add(doc.Project);
                        }
                    }
                }
                if (projects.Count == 0)
                {
                    Print("Расчеты не найдены!", KPLN_Loader.Preferences.MessageType.Error);
                    return Result.Cancelled;
                }
                Picker pp = new Picker(projects.OrderBy(x => x.Name).ToList());
                pp.ShowDialog();
                if (Picker.PickedProject != null)
                {
                    KPLN_Library_DataBase.Collections.DbProject pickedProject = Picker.PickedProject;
                    List<KPLN_Library_DataBase.Collections.DbDocument> documents = new List<KPLN_Library_DataBase.Collections.DbDocument>();
                    foreach (KPLN_Library_DataBase.Collections.DbDocument doc in KPLN_Library_DataBase.DbControll.Documents)
                    {
                        if (!IsValid(doc)) continue;
                        if (File.Exists(string.Format(@"{0}\doc_id_{1}.sqlite", MyPathes.ModelCheckerDBPath, doc.Id.ToString())))
                        {
                            if (doc.Project.Id == pickedProject.Id)
                            {
                                documents.Add(doc);
                            }
                        }
                    }
                    Picker dp = new Picker(documents.OrderBy(x => x.Department.Name).ThenBy(x => x.Name).ToList());
                    dp.ShowDialog();
                    if (Picker.PickedDocument != null)
                    {
                        KPLN_Library_DataBase.Collections.DbDocument pickedDocument = Picker.PickedDocument;
                        OutputDB form = new OutputDB(string.Format("{0}: {1}", pickedProject.Name, pickedDocument.Name), DbController.GetRows(pickedDocument.Id.ToString()), documents.OrderBy(x => x.Department.Name).ThenBy(x => x.Name).ToList());
                        form.Show();
                        return Result.Succeeded;
                    }
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }
            return Result.Failed;
        }
    }
}
