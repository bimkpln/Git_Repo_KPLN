using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Linq;

namespace KPLN_Looker.ExecutableCommand
{
    internal class ViewCloser : IExecutableCommand
    {
        private readonly ElementId _id;

        public ViewCloser(ElementId id)
        {
            _id = id;
        }
        public Result Execute(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            UIView uIView = uidoc.GetOpenUIViews().FirstOrDefault(v => v.ViewId == _id);
            uIView.Close();
            return Result.Succeeded;
        }
    }
}
