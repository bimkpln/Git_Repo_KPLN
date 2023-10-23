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
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double minZ = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            double maxZ = double.MinValue;
            foreach (Element elem in elemColl)
            {
                BoundingBoxXYZ elemBbox = PrepareElemBBox(elem);
                if (elemBbox != null)
                {
                    #region Получаю минимальную точку в каждой плоскости
                    XYZ bboxmim = elemBbox.Min;
                    double tminX = Math.Min(minX, bboxmim.X);
                    double tminY = Math.Min(minY, bboxmim.Y);
                    double tminZ = Math.Min(minZ, bboxmim.Z);
                    if (tminX < minX) minX = tminX;
                    if (tminY < minY) minY = tminY;
                    if (tminZ < minZ) minZ = tminZ;
                    #endregion

                    #region Получаю максимальную точку в каждой плоскости
                    XYZ bboxMax = elemBbox.Max;
                    double tmaxX = Math.Max(maxX, bboxMax.X);
                    double tmaxY = Math.Max(maxY, bboxMax.Y);
                    double tmaxZ = Math.Max(maxZ, bboxMax.Z);
                    if (tmaxX > maxX) maxX = tmaxX;
                    if (tmaxY > maxY) maxY = tmaxY;
                    if (tmaxZ > maxZ) maxZ = tmaxZ;
                    #endregion
                }
            }

            if (minX == minY || minX == minZ || maxX == maxY || maxX == maxZ) return null;

            return new BoundingBoxXYZ
            {
                Min = new XYZ(minX, minY, minZ) + new XYZ(-1, -1, -5),
                Max = new XYZ(maxX, maxY, maxZ) + new XYZ(1, 1, 5)
            };
        }

        /// <summary>
        /// Создание BBox для элемента
        /// </summary>
        private BoundingBoxXYZ PrepareElemBBox(Element elem)
        {
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

                BoundingBoxXYZ bbox = currentSolid.GetBoundingBox();
                Transform transform = bbox.Transform;
                BoundingBoxXYZ result = new BoundingBoxXYZ()
                {
                    Max = transform.OfPoint(bbox.Max),
                    Min = transform.OfPoint(bbox.Min),
                };

                return result;
            }

            return null;
        }
    }
}
