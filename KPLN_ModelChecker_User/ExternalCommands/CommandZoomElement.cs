using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    public class CommandZoomElement : IExecutableCommand
    {
        public CommandZoomElement(Element element, BoundingBoxXYZ box, XYZ centroid)
        {
            Element = element;
            Box = box;
            Centroid = centroid;
        }
        private Element Element { get; set; }
        private BoundingBoxXYZ Box { get; set; }
        private XYZ Centroid { get; set; }
        public Result Execute(UIApplication app)
        {
            try
            {
                CutView(Box, Centroid, app.ActiveUIDocument, Element);
                ICollection<ElementId> newSelection = new List<ElementId>() { Element.Id };
                app.ActiveUIDocument.Selection.SetElementIds(newSelection);
            }
            catch (Exception)
            {
                try
                {
                    app.ActiveUIDocument.ShowElements(Element);
                }
                catch (Exception) { }
            }
            return Result.Succeeded;
        }
        private void CutView(BoundingBoxXYZ box, XYZ centroid, UIDocument uidoc, Element element)
        {
            XYZ offsetMin = new XYZ(-5, -5, -2);
            XYZ offsetMax = new XYZ(5, 5, 1);
            View3D activeView = uidoc.ActiveView as View3D;
            ViewFamily activeViewFamily = ViewFamily.Invalid;
            ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
            activeViewFamily = viewFamilyType.ViewFamily;
            if (activeViewFamily == ViewFamily.ThreeDimensional)
            {
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Min = box.Min + new XYZ(-1, -1, -5);
                sectionBox.Max = box.Max + new XYZ(1, 1, 5);
                activeView.SetSectionBox(sectionBox);
                XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);
                activeView.SetOrientation(orientation);
                IList<UIView> views = uidoc.GetOpenUIViews();
                foreach (UIView uvView in views)
                {
                    if (uvView.ViewId.IntegerValue == activeView.Id.IntegerValue)
                    {
                        uvView.ZoomAndCenterRectangle(box.Min, box.Max);
                        break;
                    }
                }
            }
            else
            {
                uidoc.ShowElements(element);
            }
        }
        private XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
        {
            double degToRadian = Math.PI * 2 / 360;
            double angleHorizR = angleHorizD * degToRadian;
            double angleVertR = angleVertD * degToRadian;
            double a = Math.Cos(angleVertR);
            double b = Math.Cos(angleHorizR);
            double c = Math.Sin(angleHorizR);
            double d = Math.Sin(angleVertR);
            return new XYZ(a * b, a * c, d);
        }
    }
}
