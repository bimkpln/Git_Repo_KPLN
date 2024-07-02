using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Publication.ExternalCommands.PublicationSet
{
    public class CommandRemoveSet : IExecutableCommand
    {
        ViewSheetSet Set { get; set; }
        public CommandRemoveSet(ViewSheetSet set)
        {
            Set = set;
        }
        public Result Execute(UIApplication app)
        {
            //Print("Execute CommandRemoveSet", KPLN_Loader.Preferences.MessageType.System_Regular);
            Document doc = app.ActiveUIDocument.Document;
            try
            {
                using (Transaction t = new Transaction(doc, "Изменить набор листов"))
                {
                    t.Start();
                    doc.Delete(Set.Id);
                    t.Commit();
                }
                ModuleData.Form.OpenHomeTab();
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                ModuleData.Form.OpenHomeTab();
                return Result.Failed;
            }
        }
    }
}
