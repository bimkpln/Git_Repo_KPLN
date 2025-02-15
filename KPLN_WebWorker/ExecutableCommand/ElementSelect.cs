using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;

namespace KPLN_WebWorker.ExecutableCommand
{
    internal sealed class ElementSelect : IExecutableCommand
    {
        private readonly ICollection<ElementId> _elemIdColl;

        public ElementSelect(ICollection<ElementId> elemIdColl)
        {
            _elemIdColl = elemIdColl;
        }

        public Result Execute(UIApplication app)
        {
            if (app.ActiveUIDocument != null)
            {
                app.ActiveUIDocument.Selection.SetElementIds(_elemIdColl);

                return Result.Succeeded;
            }

            return Result.Failed;
        }
    }
}
