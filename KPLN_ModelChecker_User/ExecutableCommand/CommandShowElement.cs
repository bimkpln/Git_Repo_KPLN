using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandShowElement : IExecutableCommand
    {
        private readonly IEnumerable<Element> _elementCollection;

        public CommandShowElement(IEnumerable<Element> elemColl)
        {
            _elementCollection = elemColl;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Демонстрация"))
            {
                t.Start();

                app.ActiveUIDocument.ShowElements(_elementCollection.FirstOrDefault());
                app.ActiveUIDocument.Selection.SetElementIds(_elementCollection.Select(e => e.Id).ToList());

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
