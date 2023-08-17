using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Publication.ExternalCommands.PublicationSet
{
    public class CommandApplySet : IExecutableCommand
    {
        List<View> Views { get; set; }
        ViewSheetSet Set { get; set; }
        public CommandApplySet(List<View> views, ViewSheetSet set)
        {
            Views = views;
            Set = set;
        }
        public Result Execute(UIApplication app)
        {
            //Print("Execute CommandApplySet", KPLN_Loader.Preferences.MessageType.System_Regular);
            Document doc = app.ActiveUIDocument.Document;
            string name = Set.Name;
            try
            {
                using (Transaction t = new Transaction(doc, "Изменить набор листов"))
                {
                    t.Start();
                    doc.Delete(Set.Id);
                    doc.Regenerate();
                    PrintManager pm = doc.PrintManager;
                    pm.PrintRange = PrintRange.Select;
                    ViewSet newSet = new ViewSet();
                    foreach (View v in Views)
                    {
                        newSet.Insert(v);
                    }
                    ViewSheetSetting settings = pm.ViewSheetSetting;
                    settings.CurrentViewSheetSet.Views = newSet;
                    settings.SaveAs(name);
                    t.Commit();
                }
                foreach (ViewSheetSet set in new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).WhereElementIsNotElementType())
                {
                    if (set.Name == name && ModuleData.Form != null)
                    {
                        ModuleData.Form.PickSet(doc, set);
                        break;
                    }
                }
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
