using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

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
