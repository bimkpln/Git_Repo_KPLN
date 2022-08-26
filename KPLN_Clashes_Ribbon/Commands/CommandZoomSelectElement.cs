using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandZoomSelectElement : IExecutableCommand
    {
        public CommandZoomSelectElement(int id)
        {
            Id = id;
        }
        private int Id { get; }
        public Result Execute(UIApplication app)
        {
            try
            {
                if (app.ActiveUIDocument != null)
                {
                    
                    Document doc = app.ActiveUIDocument.Document;
                    Transaction t = new Transaction(doc, "Zoom");
                    t.Start();
                    Element element = doc.GetElement(new ElementId(Id));
                    ZoomTools.ZoomElement(element.get_BoundingBox(null), app.ActiveUIDocument);
                    app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId> { element.Id });
                    t.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
    }

}
