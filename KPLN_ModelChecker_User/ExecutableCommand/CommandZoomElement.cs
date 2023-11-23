using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_ModelChecker_User.ExecutableCommand
{
    internal class CommandZoomElement : IExecutableCommand
    {
        private readonly Element _element;

        private readonly IEnumerable<Element> _elementCollection;

        private readonly BoundingBoxXYZ _box;

        private readonly XYZ _centroid;

        public CommandZoomElement(Element element)
        {
            _element = element;
        }

        public CommandZoomElement(Element element, BoundingBoxXYZ box, XYZ centroid) : this(element)
        {
            _box = box;
            _centroid = centroid;
        }

        public CommandZoomElement(IEnumerable<Element> elemColl)
        {
            _elementCollection = elemColl;
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"{ModuleData.ModuleName}_Зум"))
            {
                t.Start();

                if (_elementCollection != null) 
                    PrepareAndSetView(app, _elementCollection);
                else
                    PrepareAndSetView(app, _element, _box, _centroid);
                
                t.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Подготовка и подрезка вида по коллеции элементов
        /// </summary>
        /// <returns>True, если вид удачно установлен</returns>
        private void PrepareAndSetView(UIApplication app, IEnumerable<Element> elemColl)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc.ActiveView is View3D activeView)
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
                ViewFamily activeViewFamily = viewFamilyType.ViewFamily;
                if (activeViewFamily == ViewFamily.ThreeDimensional)
                {
                    // Создание и применение подрезки
                    BoundingBoxXYZ sectionBox = PrepareElemsSumBBox(elemColl);
                    if (sectionBox == null) return;

                    activeView.SetSectionBox(sectionBox);

                    // Создание и применение ориентации вида
                    XYZ bboxCenterPnt = new XYZ(
                        (sectionBox.Max.X + sectionBox.Min.X) / 2,
                        (sectionBox.Max.Y + sectionBox.Min.Y) / 2,
                        (sectionBox.Max.Z + sectionBox.Min.Z) / 2);
                    XYZ forward_direction = VectorFromHorizVertAngles(135, -30);
                    XYZ up_direction = VectorFromHorizVertAngles(135, -30 + 90);
                    ViewOrientation3D orientation = new ViewOrientation3D(bboxCenterPnt, up_direction, forward_direction);
                    activeView.SetOrientation(orientation);

                    IList<UIView> views = uidoc.GetOpenUIViews();
                    foreach (UIView uvView in views)
                    {
                        if (uvView.ViewId.IntegerValue == activeView.Id.IntegerValue)
                        {
                            uvView.ZoomAndCenterRectangle(sectionBox.Min, sectionBox.Max);
                            app.ActiveUIDocument.Selection.SetElementIds(elemColl.Select(e => e.Id).ToList());
                            return;
                        }
                    }
                }
                else
                {
                    uidoc.ShowElements(elemColl.Select(e => e.Id).ToList());
                    app.ActiveUIDocument.Selection.SetElementIds(elemColl.Select(e => e.Id).ToList());
                    return;
                }
            }

            app.ActiveUIDocument.Selection.SetElementIds(elemColl.Select(e => e.Id).ToList());
        }

        /// <summary>
        /// Подготовка и подрезка вида по для одиночного элемента, с преднастройкой геометрии
        /// </summary>
        /// <returns>True, если вид удачно установлен</returns>
        private void PrepareAndSetView(UIApplication app, Element element, BoundingBoxXYZ box, XYZ centroid)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            if (uidoc.ActiveView is View3D activeView)
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
                ViewFamily activeViewFamily = viewFamilyType.ViewFamily;
                if (activeViewFamily == ViewFamily.ThreeDimensional)
                {
                    // Создание и применение подрезки
                    BoundingBoxXYZ sectionBox = new BoundingBoxXYZ
                    {
                        Min = box.Min + new XYZ(-1, -1, -5),
                        Max = box.Max + new XYZ(1, 1, 5)
                    };
                    activeView.SetSectionBox(sectionBox);

                    // Создание и применение ориентации вида
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
                            app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>() { _element.Id });
                            return;
                        }
                    }
                }
                else
                {
                    uidoc.ShowElements(element);
                    app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>() { _element.Id });
                    return;
                }
            }

            app.ActiveUIDocument.Selection.SetElementIds(new List<ElementId>() { _element.Id });
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

        /// <summary>
        /// Создаю максимальный BBox из набора элементов
        /// </summary>
        private BoundingBoxXYZ PrepareElemsSumBBox(IEnumerable<Element> elemColl)
        {
            List<XYZ> xyzColl = new List<XYZ>();
            foreach (Element elem in elemColl)
            {
                if (elem.Location is Location location)
                {
                    if (location is LocationPoint locPnt)
                        xyzColl.Add(locPnt.Point);
                    else if (location is LocationCurve locCrv)
                    {
                        xyzColl.Add(locCrv.Curve.GetEndPoint(0));
                        xyzColl.Add(locCrv.Curve.GetEndPoint(1));
                    }
                }
            }

            if (xyzColl.Count > 0)
            {
                XYZ minPoint = null;
                XYZ maxPoint = null;

                foreach (XYZ point in xyzColl)
                {
                    if (minPoint == null)
                    {
                        minPoint = point;
                        maxPoint = point;
                    }
                    else
                    {
                        if (point.X < minPoint.X) minPoint = new XYZ(point.X, minPoint.Y, minPoint.Z);
                        if (point.Y < minPoint.Y) minPoint = new XYZ(minPoint.X, point.Y, minPoint.Z);
                        if (point.Z < minPoint.Z) minPoint = new XYZ(minPoint.X, minPoint.Y, point.Z);

                        if (point.X > maxPoint.X) maxPoint = new XYZ(point.X, maxPoint.Y, maxPoint.Z);
                        if (point.Y > maxPoint.Y) maxPoint = new XYZ(maxPoint.X, point.Y, maxPoint.Z);
                        if (point.Z > maxPoint.Z) maxPoint = new XYZ(maxPoint.X, maxPoint.Y, point.Z);
                    }
                }

                return new BoundingBoxXYZ
                {
                    Min = minPoint + new XYZ(-1, -1, -5),
                    Max = maxPoint + new XYZ(1, 1, 5)
                };
            }

            return null;
        }

        /// <summary>
        /// Создание BBox для элемента
        /// </summary>
        private BoundingBoxXYZ PrepareElemBBox(Element elem)
        {
            BoundingBoxXYZ result = new BoundingBoxXYZ();

            Options opt = new Options() { DetailLevel = ViewDetailLevel.Fine };
            opt.ComputeReferences = true;
            GeometryElement geomElem = elem.get_Geometry(opt);

            Solid currentSolid = null;
            foreach (GeometryObject gObj in geomElem)
            {
                if (gObj is Solid solid1 && solid1.Volume != 0)
                {
                    currentSolid = solid1;
                }
                else if (gObj is GeometryInstance gInst)
                {
                    GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                    double tempVolume = 0;
                    foreach (GeometryObject gObj2 in instGeomElem)
                    {
                        if (gObj2 is Solid solid2 && solid2.Volume > tempVolume)
                        {
                            tempVolume = solid2.Volume;
                            currentSolid = solid2;
                        }
                    }
                }
                else 
                    throw new Exception($"Не удалось получить геометрию у элемента с id: {elem.Id}");

                if (currentSolid != null)
                {
                    BoundingBoxXYZ bbox = currentSolid.GetBoundingBox();
                    Transform transform = bbox.Transform;
                    result.Max = transform.OfPoint(bbox.Max);
                    result.Min = transform.OfPoint(bbox.Min);
                }

            }

            return result;
        }
    }
}
