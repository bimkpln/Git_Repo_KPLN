using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Library_DataBase;
using KPLN_Library_DataBase.Collections;
using KPLN_Library_SelectItem;
using KPLN_Library_SelectItem.Forms;
using System;
using System.Windows;
using static KPLN_Loader.Output.Output;

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
                FormSinglePick selectedProjectForm = SelectProject.CreateForm();
                bool? dialogResult = selectedProjectForm.ShowDialog();
                if (dialogResult != false)
                {
                    ReportManager mainForm = new ReportManager(selectedProjectForm.SelectedDbProject);
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
