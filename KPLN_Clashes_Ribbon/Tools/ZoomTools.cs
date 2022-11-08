using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Tools
{
    internal static class ZoomTools
    {
        public static void ZoomElement(BoundingBoxXYZ box, UIDocument uidoc)
        {
            XYZ offsetMin = new XYZ(-5, -5, -2);
            XYZ offsetMax = new XYZ(5, 5, 1);
            View3D activeView = uidoc.ActiveView as View3D;
            ViewFamily activeViewFamily = ViewFamily.Invalid;
            try
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
                activeViewFamily = viewFamilyType.ViewFamily;
            }
            catch (Exception) { }
            if (activeViewFamily == ViewFamily.ThreeDimensional)
            {
                activeView.SetSectionBox(new BoundingBoxXYZ() { Max = box.Max + offsetMax, Min = box.Min + offsetMin });
                XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                XYZ centroid = new XYZ((box.Max.X + box.Min.X) / 2, (box.Max.Y + box.Min.Y) / 2, (box.Max.Z + box.Min.Z) / 2);
                ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);
                activeView.SetOrientation(orientation);
                IList<UIView> views = uidoc.GetOpenUIViews();
                foreach (UIView uvView in views)
                {
                    if (uvView.ViewId.IntegerValue == activeView.Id.IntegerValue)
                    {
                        uvView.ZoomAndCenterRectangle(box.Min, box.Max);
                    }
                }
                return;
            }
        }
        
        public static XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
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
