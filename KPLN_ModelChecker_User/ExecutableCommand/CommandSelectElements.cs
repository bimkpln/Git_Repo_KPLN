using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandSelectElements : IExecutableCommand
    {
        private readonly IEnumerable<Element> _elementCollection;

        public CommandSelectElements(IEnumerable<Element> elemColl)
        {
            _elementCollection = elemColl.Where(el => el.IsValidObject);
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Выделить"))
            {
                t.Start();

                app.ActiveUIDocument.Selection.SetElementIds(_elementCollection.Select(e => e.Id).ToList());

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
