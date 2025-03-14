using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_TaskManager.ExecutableCommand
{
    internal class CommandShowElement : IExecutableCommand
    {
        private readonly ICollection<ElementId> _elemIdColl;

        public CommandShowElement(ICollection<ElementId> elemIdColl)
        {
            _elemIdColl = elemIdColl;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Демонстрация"))
            {
                t.Start();

                app.ActiveUIDocument.ShowElements(_elemIdColl.FirstOrDefault());
                app.ActiveUIDocument.Selection.SetElementIds(_elemIdColl);

                t.Commit();
            }

            return Result.Succeeded;
        }
    }
}
