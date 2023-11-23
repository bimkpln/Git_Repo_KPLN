using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.Common
{
    internal class CheckLevelOfInstanceData
    {
        private readonly BuiltInParameter[] _sysElemBicDownOffsetParams = new BuiltInParameter[]
        {
            BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM,
            BuiltInParameter.WALL_BASE_OFFSET,
            BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM,
            BuiltInParameter.STAIRS_RAILING_HEIGHT_OFFSET,
            BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,
        };

        private readonly BuiltInParameter[] _sysElemBicUpOffsetParams = new BuiltInParameter[]
        {
            BuiltInParameter.WALL_TOP_OFFSET,
        };

        private readonly BuiltInParameter[] _userElemBicDownOffsetParams = new BuiltInParameter[]
        {
            BuiltInParameter.INSTANCE_ELEVATION_PARAM,
            //BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
            //BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,
        };

        private readonly BuiltInParameter[] _userElemBicUpOffsetParams = new BuiltInParameter[]
        {
            BuiltInParameter.WALL_TOP_OFFSET,
        };

        private List<BoundingBoxXYZ> _currentBBoxColl = new List<BoundingBoxXYZ>();
        private List<Solid> _currentSolidColl = new List<Solid>();
        private Level _currentElemProjectDownLevel;
        private Level _currentElemProjectUpLevel;

        public CheckLevelOfInstanceData(Element elem)
        {
            CurrentElem = elem;
        }

        public Element CurrentElem { get; }

        public Element CurrentElemHostElem { get; }

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
        /// Уровень элемента в проекте - привязка снизу
        /// </summary>
        public Level CurrentElemProjectDownLevel
        {
            get => _currentElemProjectDownLevel;
            private set => _currentElemProjectDownLevel = value;
        }

        /// <summary>
        /// Уровень элемента в проекте - привязка сверху
        /// </summary>
        public Level CurrentElemProjectUpLevel
        {
            get => _currentElemProjectUpLevel;
            private set => _currentElemProjectUpLevel = value;
        }

        /// <summary>
        /// Отступ от уровня вверх
        /// </summary>
        public double UpOffset { get; private set; }
        
        /// <summary>
        /// Отступ от уровня вниз
        /// </summary>
        public double DownOffset { get; private set; }

        /// <summary>
        /// Метка, указывающая статус анализа у элемента. По умолчанию - true, т.е. проверка НЕ проводилась
        /// </summary>
        public bool IsEmptyChecked { get; set; } = true;

        /// <summary>
        /// Инициализация коллекции Solidов для элемента
        /// </summary>
        public CheckLevelOfInstanceData SetCurrentSolidColl()
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
        public CheckLevelOfInstanceData SetCurrentBBoxColl()
        {
            foreach (Solid solid in CurrentSolidColl)
            {
                CurrentBBoxColl.Add(GetBoundingBoxXYZ(solid));
            }

            if (CurrentBBoxColl.Count == 0)
                throw new Exception($"Не удалось получить BoundingBoxXYZ у элемента с id: {CurrentElem.Id}");

            return this;
        }

        /// <summary>
        /// Инициализация проектного уровня для элемента
        /// </summary>
        public CheckLevelOfInstanceData SetCurrentProjectLevel()
        {
            Document doc = CurrentElem.Document;
            if (CurrentElem is Wall wall)
                CurrentElemProjectUpLevel = doc.GetElement(wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).AsElementId()) as Level;
            
            if (doc.GetElement(CurrentElem.LevelId) is Level tempLevel)
                CurrentElemProjectDownLevel = tempLevel;
            else
            {
                if (CurrentElem is FamilyInstance famInst)
                {
                    // На основе линии
                    if (famInst.Host != null && doc.GetElement(famInst.Host.Id) is Level hostLevel)
                        CurrentElemProjectDownLevel = hostLevel;
                    // На основе элемента
                    else if (famInst.Host != null && doc.GetElement(famInst.Host.LevelId) is Level tempHostLevel)
                        CurrentElemProjectDownLevel = tempHostLevel;
                    // Остальные основы
                    else
                    {
                        // На основе грани
                        Parameter schedParam = CurrentElem.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
                        if (schedParam != null)
                            CurrentElemProjectDownLevel = doc.GetElement(schedParam.AsElementId()) as Level;
                    }
                }

            }

            return this;
        }

        /// <summary>
        /// Инициализация отсутпов от уровня
        /// </summary>
        public CheckLevelOfInstanceData SetOffsets()
        {
            if (CurrentElem is FamilyInstance _)
            {
                SetDownOffset(GetOffsets(_userElemBicDownOffsetParams));
                SetUpOffset(GetOffsets(_userElemBicUpOffsetParams));
            }
            else
            {
                SetDownOffset(GetOffsets(_sysElemBicDownOffsetParams));
                SetUpOffset(GetOffsets(_sysElemBicUpOffsetParams));
            }

            return this;
        }

        private List<double> GetOffsets(BuiltInParameter[] _bicOffsetParams)
        {
            List<double> offsets = new List<double>();
            foreach (BuiltInParameter bic in _bicOffsetParams)
            {
                Parameter param = CurrentElem.get_Parameter(bic);
                if (param != null)
                {
                    offsets.Add(Math.Round(param.AsDouble(), 5));
                }
            }

            return offsets;
        }

        private void SetDownOffset(List<double> downOffsets)
        {
            if (!downOffsets.Any())
                throw new Exception($"Не удалось определить отметку id: {CurrentElem.Id}");
            else if (downOffsets.All(o => o == 0))
                DownOffset = 0;
            else
                DownOffset = downOffsets.Min();
        }

        private void SetUpOffset(List<double> upOffsets)
        {
            if (!upOffsets.Any() || upOffsets.All(o => o == 0))
                UpOffset = 0;
            else
                UpOffset = upOffsets.Max();
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
