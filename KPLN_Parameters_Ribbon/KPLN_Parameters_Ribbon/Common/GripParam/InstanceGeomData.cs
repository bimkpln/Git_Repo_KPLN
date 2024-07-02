using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Контейнер для хранения геометрических данных для InstanceElemData
    /// </summary>
    internal class InstanceGeomData : InstanceElemData
    {
        private readonly List<XYZ> _currentGeomCenterColl = new List<XYZ>();
        private List<BoundingBoxXYZ> _currentBBoxColl = new List<BoundingBoxXYZ>();
        private List<Solid> _currentSolidColl = new List<Solid>();
        private double _sumSolidVolume = 0;
        private double[] _minAndMaxElevation;

        public InstanceGeomData(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Коллекция Solid элемента
        /// </summary>
        public List<Solid> CurrentSolidColl
        {
            get => _currentSolidColl;
            private set => _currentSolidColl = value;
        }

        /// <summary>
        /// Коллекция BoundingBoxXYZ элемента
        /// </summary>
        public List<BoundingBoxXYZ> CurrentBBoxColl
        {
            get => _currentBBoxColl;
            private set => _currentBBoxColl = value;
        }

        /// <summary>
        /// Коллекция точек геом. центров составного элемента
        /// </summary>
        public List<XYZ> CurrentGeomCenterColl
        {
            get
            {
                if (!_currentGeomCenterColl.Any())
                {
                    List<XYZ> tempColl = new List<XYZ>();
                    try
                    {
                        foreach (Solid solid in CurrentSolidColl)
                        {
                            tempColl.Add(solid.ComputeCentroid());
                        }
                    }
                    // Для сожной геометрии (разуклонка больших перекрытий) - могут быть проблемы с центроидом
                    catch
                    {
                        tempColl.Clear();
                        foreach (BoundingBoxXYZ instBbox in CurrentBBoxColl)
                        {
                            tempColl.Add(0.5 * (instBbox.Max + instBbox.Min));
                        }
                    }
                    _currentGeomCenterColl.AddRange(tempColl);
                }

                return _currentGeomCenterColl;
            }
        }

        /// <summary>
        /// Минимальная и максимальная отметки элемента
        /// </summary>
        public double[] MinAndMaxElevation
        {
            get
            {
                if (_minAndMaxElevation == null)
                {
                    _minAndMaxElevation = new double[2];

                    double minElevOfElem = double.MaxValue;
                    double maxElevOfElem = double.MinValue;
                    foreach (BoundingBoxXYZ instBbox in CurrentBBoxColl)
                    {
                        double minZ = instBbox.Min.Z;
                        double maxZ = instBbox.Max.Z;
                        if (minZ < minElevOfElem)
                            minElevOfElem = minZ;
                        if (maxZ > maxElevOfElem)
                            maxElevOfElem = maxZ;
                    }

                    _minAndMaxElevation[0] = minElevOfElem;
                    _minAndMaxElevation[1] = maxElevOfElem;
                }

                return _minAndMaxElevation;
            }
        }

        public double SumSolidVolume
        {
            get
            {
                if (_sumSolidVolume == 0)
                {
                    _sumSolidVolume = CurrentSolidColl
                        .Where(s => s.Volume > 0)
                        .Select(s => s.Volume)
                        .Aggregate((a, b) => a + b);
                }

                return _sumSolidVolume;
            }
        }

        /// <summary>
        /// Инициализация коллекции Solidов для элемента
        /// </summary>
        public InstanceGeomData SetCurrentSolidColl()
        {
            Options opt = new Options() { DetailLevel = ViewDetailLevel.Fine };
            opt.ComputeReferences = true;
            GeometryElement geomElem = CurrentElem.get_Geometry(opt);
            if (geomElem != null)
            {
                List<Solid> solidColl = new List<Solid>();
                GetSolidsFromGeomElem(geomElem, Transform.Identity, solidColl);
                if (solidColl.All(s => s.Volume == 0))
                {
                    BoundingBoxXYZ bbox = CurrentElem.get_BoundingBox(null);
                    XYZ bboxMin = bbox.Min;
                    XYZ bboxMax = bbox.Max;

                    #region Уточнение координат по плоским эл-м
                    // Очистка от плоских эл-в (по оси X)
                    if (Math.Round(bboxMin.X, 5) == Math.Round(bboxMax.X, 5))
                    {
                        double newMinX = bboxMin.X > 0 ? bboxMin.X - 0.1 : bboxMin.X + 0.1;
                        double newMaxX = bboxMax.X > 0 ? bboxMax.X + 0.1 : bboxMax.X - 0.1;
                        bboxMin = new XYZ(newMinX, bboxMin.Y, bboxMin.Z);
                        bboxMax = new XYZ(newMaxX, bboxMax.Y, bboxMax.Z);
                    }
                    // Очистка от плоских эл-в (по оси Y)
                    if (Math.Round(bboxMin.Y, 5) == Math.Round(bboxMax.Y, 5))
                    {
                        double newMinY = bboxMin.Y > 0 ? bboxMin.Y - 0.1 : bboxMin.Y + 0.1;
                        double newMaxY = bboxMax.Y > 0 ? bboxMax.Y + 0.1 : bboxMax.Y - 0.1;
                        bboxMin = new XYZ(bboxMin.X, newMinY, bboxMin.Z);
                        bboxMax = new XYZ(bboxMax.X, newMaxY, bboxMax.Z);
                    }
                    // Очистка от плоских эл-в (по оси Z)
                    if (Math.Round(bboxMin.Z, 5) == Math.Round(bboxMax.Z, 5))
                        bboxMax = new XYZ(bboxMax.X, bboxMax.Y, bboxMin.Z + 0.1);
                    #endregion

                    List<XYZ> pointsDwn = new List<XYZ>()
                    {
                        new XYZ(bboxMin.X, bboxMin.Y, bboxMin.Z),
                        new XYZ(bboxMin.X, bboxMax.Y, bboxMin.Z),
                        new XYZ(bboxMax.X, bboxMax.Y, bboxMin.Z),
                        new XYZ(bboxMax.X, bboxMin.Y, bboxMin.Z),
                    };
                    List<XYZ> pointsUp = new List<XYZ>()
                    {
                        new XYZ(bboxMin.X, bboxMin.Y, bboxMax.Z),
                        new XYZ(bboxMin.X, bboxMax.Y, bboxMax.Z),
                        new XYZ(bboxMax.X, bboxMax.Y, bboxMax.Z),
                        new XYZ(bboxMax.X, bboxMin.Y, bboxMax.Z),
                    };


                    List<Curve> curvesListDwn = GetCurvesListFromPoints(pointsDwn);
                    List<Curve> curvesListUp = GetCurvesListFromPoints(pointsUp);
                    CurveLoop curveLoopDwn = CurveLoop.Create(curvesListDwn);
                    CurveLoop curveLoopUp = CurveLoop.Create(curvesListUp);

                    CurveLoop[] curves = new CurveLoop[] { curveLoopDwn, curveLoopUp };
                    SolidOptions solidOptions = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                    solidColl.Add(GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions));
                }

                CurrentSolidColl.AddRange(solidColl.Where(s => s.Volume != 0));
            }

            return this;
        }

        /// <summary>
        /// Инициализация коллекции BoundingBoxXYZов для элемента
        /// </summary>
        public InstanceGeomData SetCurrentBBoxColl()
        {
            foreach (Solid solid in CurrentSolidColl)
            {
                CurrentBBoxColl.Add(GetBoundingBoxXYZ(solid));
            }

            if (CurrentBBoxColl.Count == 0)
                throw new Exception($"Не удалось получить BoundingBoxXYZ у элемента с id: {CurrentElem.Id}");

            return this;
        }

        private List<Curve> GetCurvesListFromPoints(List<XYZ> pointsOfIntersect)
        {
            List<Curve> curvesList = new List<Curve>();
            for (int i = 0; i < pointsOfIntersect.Count; i++)
            {
                if (i == pointsOfIntersect.Count - 1)
                {
                    curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[0]));
                    continue;
                }

                curvesList.Add(Line.CreateBound(pointsOfIntersect[i], pointsOfIntersect[i + 1]));
            }
            return curvesList;
        }

        /// <summary>
        /// Получить Solid из элемента
        /// </summary>
        private void GetSolidsFromGeomElem(GeometryElement geometryElement, Transform transformation, IList<Solid> solids)
        {
            foreach (GeometryObject geomObject in geometryElement)
            {
                switch (geomObject)
                {
                    case Solid solid:
                        solids.Add(solid);
                        break;

                    case GeometryInstance geomInstance:
                        GetSolidsFromGeomElem(geomInstance.GetInstanceGeometry(), geomInstance.Transform.Multiply(transformation), solids);
                        break;

                    case GeometryElement geomElem:
                        GetSolidsFromGeomElem(geomElem, transformation, solids);
                        break;
                }
            }
        }

        /// <summary>
        /// Получить BoundingBoxXYZ из Solid
        /// </summary>
        private BoundingBoxXYZ GetBoundingBoxXYZ(Solid solid)
        {
            if (solid != null && solid.Volume != 0)
            {
                BoundingBoxXYZ bbox = solid.GetBoundingBox();
                Transform transform = bbox.Transform;
                Transform resultTrans = transform;
                return new BoundingBoxXYZ()
                {
                    Max = resultTrans.OfPoint(bbox.Max),
                    Min = resultTrans.OfPoint(bbox.Min),
                };
            }

            return null;
        }
    }
}
