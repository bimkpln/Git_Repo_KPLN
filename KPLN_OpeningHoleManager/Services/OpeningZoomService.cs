using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Логика секущего бокса и ракурса продублирована из KPLN_Clashes_Ribbon.Tools.ZoomTools.
    /// Вместо ClashPoint приближаемся к выбранному отверстию.
    /// </summary>
    internal static class OpeningZoomService
    {
        internal static void Show(UIDocument uidoc, Element opening)
        {
            Document doc = uidoc.Document;
            BoundingBoxXYZ box = opening.get_BoundingBox(null) ?? throw new InvalidOperationException("У отверстия нет габаритов для приближения.");

            box = ToModelBounds(box);
            View3D view = uidoc.ActiveView as View3D;
            if (!CanZoom(view))
            {
                view = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>()
                    .FirstOrDefault(CanZoom) ?? throw new InvalidOperationException("Откройте незаблокированный ортогональный 3D-вид для показа отверстия.");
                uidoc.ActiveView = view;
            }


            // Отступы и ракурс совпадают с приближением к ClashPoint.
            OpeningZoomSettings settings = LoadSettings();
            XYZ offsetMin = new XYZ(
                -ToFeet(settings.OffsetXMillimeters, OpeningZoomSettings.DefaultOffsetXYMillimeters),
                -ToFeet(settings.OffsetYMillimeters, OpeningZoomSettings.DefaultOffsetXYMillimeters),
                -ToFeet(settings.OffsetZMinMillimeters, OpeningZoomSettings.DefaultOffsetZMinMillimeters));
            XYZ offsetMax = new XYZ(
                ToFeet(settings.OffsetXMillimeters, OpeningZoomSettings.DefaultOffsetXYMillimeters),
                ToFeet(settings.OffsetYMillimeters, OpeningZoomSettings.DefaultOffsetXYMillimeters),
                ToFeet(settings.OffsetZMaxMillimeters, OpeningZoomSettings.DefaultOffsetZMaxMillimeters));
            BoundingBoxXYZ zoomBox = new BoundingBoxXYZ { Min = box.Min + offsetMin, Max = box.Max + offsetMax };
            XYZ forwardDirection = VectorFromHorizVertAngles(135, -30);
            XYZ upDirection = VectorFromHorizVertAngles(135, -30 + 90);
            XYZ centroid = new XYZ((box.Max.X + box.Min.X) / 2, (box.Max.Y + box.Min.Y) / 2, (box.Max.Z + box.Min.Z) / 2);
            ViewOrientation3D orientation = new ViewOrientation3D(centroid, upDirection, forwardDirection);

            using (Transaction tx = new Transaction(doc, "KPLN: Показать отверстие"))
            {
                tx.Start();
                view.IsSectionBoxActive = true;
                view.SetSectionBox(zoomBox);
                view.SetOrientation(orientation);
                if (tx.Commit() != TransactionStatus.Committed)
                    throw new InvalidOperationException("Revit отменил настройку вида для показа отверстия.");
            }


            // Зумируем после фиксации бокса, чтобы обновление вида не сбросило приближение.
            uidoc.Selection.SetElementIds(new List<ElementId> { opening.Id });
            uidoc.RefreshActiveView();
            UIView uiView = uidoc.GetOpenUIViews().FirstOrDefault(candidate => candidate.ViewId.Equals(view.Id))
                ?? throw new InvalidOperationException("Не найдено открытое окно 3D-вида для приближения.");

            uiView.ZoomAndCenterRectangle(zoomBox.Min, zoomBox.Max);
        }

        private static bool CanZoom(View3D view) =>
            view != null && !view.IsTemplate && !view.IsPerspective && !view.IsLocked;

        private static OpeningZoomSettings LoadSettings()
        {
            try
            {
                return ConfigService.ReadConfigFile<OpeningZoomSettings>(ConfigType.Local, "Clashes") as OpeningZoomSettings
                    ?? new OpeningZoomSettings();
            }
            catch (Exception)
            {
                return new OpeningZoomSettings();
            }
        }

        private static double ToFeet(double millimeters, double fallback)
        {
            if (double.IsNaN(millimeters) || double.IsInfinity(millimeters) || millimeters < 0)
                millimeters = fallback;

            return millimeters / OpeningZoomSettings.FeetToMillimeters;
        }

        private static BoundingBoxXYZ ToModelBounds(BoundingBoxXYZ box)
        {
            var points = new List<XYZ>();
            foreach (double x in new[] { box.Min.X, box.Max.X })
                foreach (double y in new[] { box.Min.Y, box.Max.Y })
                    foreach (double z in new[] { box.Min.Z, box.Max.Z })
                        points.Add(box.Transform.OfPoint(new XYZ(x, y, z)));

            return new BoundingBoxXYZ
            {
                Min = new XYZ(points.Min(p => p.X), points.Min(p => p.Y), points.Min(p => p.Z)),
                Max = new XYZ(points.Max(p => p.X), points.Max(p => p.Y), points.Max(p => p.Z)),
            };
        }

        private static XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
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
