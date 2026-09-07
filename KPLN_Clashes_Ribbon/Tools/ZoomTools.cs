using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Clashes_Ribbon.Core;
using KPLN_Clashes_Ribbon.Services;
using System;
using System.Collections.Generic;

namespace KPLN_Clashes_Ribbon.Tools
{
    internal static class ZoomTools
    {
        public static void ZoomElement(BoundingBoxXYZ box, UIDocument uidoc, View3D activeView, BoundingBoxXYZ orientationBox = null)
        {
            if (box == null || uidoc == null || activeView == null)
                return;

            ZoomSettings settings = ZoomSettingsConfigService.LoadOrCreateDefault(uidoc.Document);
            XYZ offsetMin = new XYZ(-settings.OffsetXFeet, -settings.OffsetYFeet, -settings.OffsetZMinFeet);
            XYZ offsetMax = new XYZ(settings.OffsetXFeet, settings.OffsetYFeet, settings.OffsetZMaxFeet);

            ViewFamily activeViewFamily = ViewFamily.Invalid;
            try
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
                activeViewFamily = viewFamilyType.ViewFamily;
            }
            catch (Exception) { }

            if (activeViewFamily == ViewFamily.ThreeDimensional)
            {
                BoundingBoxXYZ zoomBox = new BoundingBoxXYZ() { Max = box.Max + offsetMax, Min = box.Min + offsetMin };
                activeView.SetSectionBox(zoomBox);

                BoundingBoxXYZ centroidBox = orientationBox ?? box;
                XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                XYZ centroid = new XYZ((centroidBox.Max.X + centroidBox.Min.X) / 2, (centroidBox.Max.Y + centroidBox.Min.Y) / 2, (centroidBox.Max.Z + centroidBox.Min.Z) / 2);
                ViewOrientation3D orientation = new ViewOrientation3D(centroid, up_direction, forward_direction);

                activeView.SetOrientation(orientation);

                IList<UIView> views = uidoc.GetOpenUIViews();
                foreach (UIView uvView in views)
                {
                    if (uvView.ViewId.Equals(activeView.Id))
                        uvView.ZoomAndCenterRectangle(zoomBox.Min, zoomBox.Max);
                }
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
