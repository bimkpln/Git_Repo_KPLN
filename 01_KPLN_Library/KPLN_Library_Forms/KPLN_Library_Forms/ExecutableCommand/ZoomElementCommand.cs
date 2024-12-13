using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Library_Forms.ExecutableCommand
{
    /// <summary>
    /// Комманда для зуммирования элемента (открытие вида, выделение и т.п.)
    /// </summary>
    public class ZoomElementCommand : IExecutableCommand
    {
        private readonly Element _element;
        private readonly IEnumerable<Element> _elementCollection;

        private BoundingBoxXYZ _box;
        private XYZ _centroid;

        /// <summary>
        /// По элементу
        /// </summary>
        public ZoomElementCommand(Element element)
        {
            _element = element;
        }

        /// <summary>
        /// По элементу из связи
        /// </summary>
        public ZoomElementCommand(Element element, RevitLinkInstance linkInstance) : this(element)
        {
            CurrentLinkInstance = linkInstance;
        }

        /// <summary>
        /// По элементу и предварительно подготовленному боксу
        /// </summary>
        public ZoomElementCommand(Element element, BoundingBoxXYZ box, XYZ centroid) : this(element)
        {
            ZoomBBox = box;
            ZoomCentroid = centroid;
        }

        /// <summary>
        /// По коллекции элементов
        /// </summary>
        public ZoomElementCommand(IEnumerable<Element> elemColl)
        {
            _elementCollection = elemColl;
        }

        internal BoundingBoxXYZ ZoomBBox
        {
            get
            {
                if (_box == null)
                {
                    if (_element != null)
                        _box = PrepareElemBBox(_element);
                    else
                        _box = PrepareElemsSumBBox(_elementCollection);
                }

                return _box;
            }

            private set { _box = value; }
        }

        internal XYZ ZoomCentroid
        {
            get
            {
                if (_centroid == null)
                    _centroid = new XYZ((ZoomBBox.Min.X + ZoomBBox.Max.X) / 2, (ZoomBBox.Min.Y + ZoomBBox.Max.Y) / 2, (ZoomBBox.Min.Z + ZoomBBox.Max.Z) / 2);

                return _centroid;
            }

            private set { _centroid = value; }
        }

        internal RevitLinkInstance CurrentLinkInstance { get; } = null;

        public Transform CurrentLinkTransform
        {
            get
            {
                if (CurrentLinkInstance != null)
                    return CurrentLinkInstance.GetTotalTransform();
                
                return null;
            }
        }

        public Result Execute(UIApplication app)
        {
            using (Transaction t = new Transaction(app.ActiveUIDocument.Document, $"KPLN: Зум"))
            {
                t.Start();

                if (_elementCollection != null)
                    PrepareAndSetView(app, _elementCollection, ZoomBBox);
                else
                    PrepareAndSetView(app, new List<Element>() { _element }, ZoomBBox);

                t.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Подготовка и подрезка вида по коллеции элементов
        /// </summary>
        private void PrepareAndSetView(UIApplication app, IEnumerable<Element> elemColl, BoundingBoxXYZ sectionBox)
        {
            if (sectionBox == null) 
                return;
            
            app.ActiveUIDocument.Selection.SetElementIds(elemColl.Select(e => e.Id).ToList());

            UIDocument uidoc = app.ActiveUIDocument;
            // Подрезка при открытом 3d-виде
            if (uidoc.ActiveView is View3D activeView)
            {
                ViewFamilyType viewFamilyType = uidoc.Document.GetElement(activeView.GetTypeId()) as ViewFamilyType;
                ViewFamily activeViewFamily = viewFamilyType.ViewFamily;
                if (activeViewFamily == ViewFamily.ThreeDimensional)
                {
                    // Создание и применение подрезки
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

                    UIView uiView = uidoc
                        .GetOpenUIViews()
                        .Where(view => view.ViewId.IntegerValue == activeView.Id.IntegerValue)
                        .FirstOrDefault();
                    uiView.ZoomAndCenterRectangle(sectionBox.Min, sectionBox.Max);
                    return;
                }
                else
                {
                    uidoc.ShowElements(elemColl.Select(e => e.Id).ToList());
                    return;
                }
            }

            // Вызов команды SelectionBox для спек и других таблиц
            else if (uidoc.ActiveView is TableView)
            {
                RevitCommandId selBoxommId = RevitCommandId.LookupPostableCommandId(PostableCommand.SelectionBox);
                if (app.CanPostCommand(selBoxommId))
                    app.PostCommand(selBoxommId);

                return;
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
        /// Создание BBox для элемента (РЕСУРСОЕМКИЙ МЕТОД)
        /// </summary>
        private BoundingBoxXYZ PrepareElemBBox(Element elem)
        {
            #region Задаю Solid
            Solid resultSolid = null;
            Options opt = new Options
            {
                DetailLevel = ViewDetailLevel.Fine,
                ComputeReferences = true
            };

            Solid tempSolid = null;
            GeometryElement geomElem = elem.get_Geometry(opt);
            foreach (GeometryObject gObj in geomElem)
            {
                Solid solid = gObj as Solid;
                GeometryInstance gInst = gObj as GeometryInstance;
                if (solid != null) tempSolid = solid;
                else if (gInst != null)
                {
                    GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                    double tempVolume = 0;
                    foreach (GeometryObject gObj2 in instGeomElem)
                    {
                        solid = gObj2 as Solid;
                        if (solid != null && solid.Volume > tempVolume)
                        {
                            tempVolume = solid.Volume;
                            tempSolid = solid;
                        }
                    }
                }

                resultSolid = CurrentLinkInstance == null ? tempSolid : SolidUtils.CreateTransformed(resultSolid, CurrentLinkTransform);
            }
            #endregion

            #region Задаю BoundingBoxXYZ
            if (resultSolid != null)
            {
                BoundingBoxXYZ bbox = resultSolid.GetBoundingBox() ?? throw new Exception($"Элементу {elem.Id} - невозможно создать BoundingBoxXYZ. Отправь сообщение разработчику");
                Transform transform = bbox.Transform;
                Transform resultTransform = CurrentLinkInstance == null ? transform : transform * CurrentLinkTransform;

                return new BoundingBoxXYZ()
                {
                    Max = resultTransform.OfPoint(bbox.Max) + new XYZ(-1, -1, -5),
                    Min = resultTransform.OfPoint(bbox.Min) + new XYZ(1, 1, 5),
                };
            }
            #endregion

            return null;
        }
    }
}
