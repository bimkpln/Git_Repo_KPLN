using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Publication.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Publication.ExternalCommands.PublicationSet
{
    public class CommandCreateSet : IExecutableCommand
    {
        List<View> Views { get; set; }
        private string Name { get; set; }
        public CommandCreateSet(List<View> views, string name)
        {
            Name = name;
            Views = views;
        }
        public Result Execute(UIApplication app)
        {
            ModuleData.Form.OpenWaitTab();
            Document doc = app.ActiveUIDocument.Document;
            try
            {
                using ( Transaction t = new Transaction(doc, "Изменить набор листов"))
                {
                    t.Start();
                    PrintManager pm = doc.PrintManager;
                    pm.PrintRange = PrintRange.Select;
                    ViewSet newSet = new ViewSet();
                    foreach (View v in Views)
                    {
                        newSet.Insert(v);
                    }
                    ViewSheetSetting settings = pm.ViewSheetSetting;
                    settings.CurrentViewSheetSet.Views = newSet;
                    settings.SaveAs(Name);
                    t.Commit();
                }
                foreach (ViewSheetSet set in new FilteredElementCollector(doc).OfClass(typeof(ViewSheetSet)).WhereElementIsNotElementType())
                {
                    if (set.Name == Name && ModuleData.Form != null)
                    {
                        ModuleData.Form.PickSet(doc, set);
                        break;
                    }
                }
                ModuleData.Form.OpenHomeTab();
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                ModuleData.Form.Hide();
                TaskDialog TD = new TaskDialog("Ошибка");
                TD.TitleAutoPrefix = false;
                TD.MainContent = "Имя занято! Попробуйте создать набор с другим именем.";
                TD.Show();
                ModuleData.Form.OpenHomeTab();
                ModuleData.Form.Show();
                return Result.Failed;
            }
        }
    }
}
