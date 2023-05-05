using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandShowElement : IExecutableCommand
    {
        private Element _element;

        public CommandShowElement(Element element)
        {
            _element = element;
        }

        public Result Execute(UIApplication app)
        {
            app.ActiveUIDocument.ShowElements(_element);
            app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>() { _element.Id });

            return Result.Succeeded;
        }
    }
}
