using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Общий класс для обработки элементов с построением солидов
    /// </summary>
    internal class GripByGeometry
    {
        private readonly Document _doc;

        private readonly string _levelParamName;

        private readonly string _sectionParamName;

        public int PbCounter = 0;

        private List<DirectShape> _directShapesColl = new List<DirectShape>();

        private Dictionary<Element, List<string>> _duplicatesWriteParamElems = new Dictionary<Element, List<string>>();

        /// <summary>
        /// Словарь элементов, которые подверглись повторной записи параметров
        /// </summary>
        public IReadOnlyDictionary<Element, List<string>> DuplicatesWriteParamElems
        {
            get { return _duplicatesWriteParamElems; }
        }

        public GripByGeometry(Document doc, string levelParamName, string sectionParamName)
        {
            _doc = doc;
            _sectionParamName = sectionParamName;
            _levelParamName = levelParamName;
        }

        /// <summary>
        /// Анализ уровней (получение нижнего минимального и верхнего)
        /// </summary>
        /// <returns></returns>
        public IEnumerable<MyLevel> LevelPrepare()
        {
            List<MyLevel> preapareLevels = new List<MyLevel>();

            IEnumerable<Level> levelColl = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>()
                    .OrderBy(x => x.Elevation);

            IEnumerable<Level> equalAboveLevelColl = new List<Level>();
            foreach (Level level in levelColl)
            {
                if (equalAboveLevelColl.Contains(level, new ElementComparerTool())) { continue; }

                ElementId aboveLevelId = level.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId();
                Level aboveLevel = (Level)_doc.GetElement(aboveLevelId);

                equalAboveLevelColl = levelColl.Where(x => x.get_Parameter(BuiltInParameter.LEVEL_UP_TO_LEVEL).AsElementId() == aboveLevelId);

                Level minElevLevel = equalAboveLevelColl.OrderBy(x => x.Elevation).FirstOrDefault();

                MyLevel myLevel = new MyLevel(minElevLevel, aboveLevel);

                preapareLevels.Add(myLevel);
            }

            return preapareLevels;
        }

        /// <summary>
        /// Анализ осей проекта
        /// </summary>
        /// <param name="gridSectionParam">Параметр осей, откуда берется номер секции</param>
        /// <returns>Коллекция осей, которые являются граничными для секций</returns>
        /// <exception cref="Exception"></exception>
        public Dictionary<string, HashSet<Grid>> GridPrepare(string gridSectionParam)
        {
            Dictionary<string, HashSet<Grid>> sectionsGrids = new Dictionary<string, HashSet<Grid>>();

            List<Grid> grids = new FilteredElementCollector(_doc).WhereElementIsNotElementType().OfClass(typeof(Grid)).Cast<Grid>().ToList();
            foreach (Grid grid in grids)
            {
                Parameter param = grid.LookupParameter(gridSectionParam);
                if (param == null) continue;

                string valueOfParam = param.AsString();
                if (valueOfParam == null || valueOfParam.Length == 0) continue;

                foreach (string item in valueOfParam.Split('-'))
                {
                    if (sectionsGrids.ContainsKey(item))
                    {
                        sectionsGrids[item].Add(grid);
                        continue;
                    }
                    sectionsGrids.Add(item, new HashSet<Grid>() { grid });
                }

            }

            foreach (string sg in sectionsGrids.Keys)
            {
                if (sectionsGrids[sg].Count < 4)
                {
                    throw new Exception($"Количество осей с номером секции: {sg} меньше 4. Проверьте назначение параметров у осей!");
                }
            }

            if (sectionsGrids.Keys.Count == 0)
            {
                throw new Exception($"Для заполнения номера секции в элементах, необходимо заполнить параметр: {gridSectionParam} в осях! Значение указывается через \"-\" для осей, относящихся к нескольким секциям.");
            }

            return sectionsGrids;
        }

        /// <summary>
        /// Создание внутренней коллекции солидов с атрибутами
        /// </summary>
        /// <param name="sectionsGrids">Коллекция граничных осей для солида</param>
        /// <param name="floorTextPosition">Позиция индекса уровня</param>
        /// <param name="splitChar">Разделитель в именах уровней</param>
        /// <returns>Коллекция солидов, ограниченных сетками (оси и уровни) Ревит</returns>
        public IEnumerable<MySolid> SolidsCollectionPrepare(
            Dictionary<string, HashSet<Grid>> sectionsGrids,
            IEnumerable<MyLevel> myLevels,
            int floorTextPosition,
            char splitChar)

        {
            List<MySolid> mySolidsColl = new List<MySolid>();

            foreach (MyLevel myLevel in myLevels)
            {
                double[] minMaxCoords = GetMinMaxZCoordOfLevel(myLevel.CurrentLevel, myLevel.AboveLevel);

                string levelIndex = GetFloorNumberByLevel(myLevel.CurrentLevel, floorTextPosition, splitChar);

                //Анализ секций
                foreach (string sectIndex in sectionsGrids.Keys)
                {
                    List<Grid> gridsOfSect = sectionsGrids[sectIndex].ToList();
                    List<XYZ> pointsOfGridsIntersect = GetPointsOfGridsIntersection(gridsOfSect);
                    pointsOfGridsIntersect.Sort(new ClockwiseComparerTool(GetCenterPointOfPoints(pointsOfGridsIntersect)));

                    _directShapesColl.Add(CreateSolidsInModel(minMaxCoords, pointsOfGridsIntersect, out Solid solid));

                    MySolid mySolid = new MySolid(solid, levelIndex, sectIndex);

                    mySolidsColl.Add(mySolid);
                }
            }

            return mySolidsColl;
        }

        /// <summary>
        /// Поиск элементов, которые пересекаются с коллекцией солидов
        /// </summary>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="mySolidsColl">Коллекция спец. солидов, для анализа на вхождение</param>
        /// <returns>Коллекция элементов, которые не попали в пересечения с солидами</returns>
        public IEnumerable<Element> IntersectWithSolidExcecute(
            List<Element> elems,
            IEnumerable<MySolid> mySolidsColl,
            Progress_Single pb)

        {
            IEnumerable<Element> notIntersectedElems = elems;

            foreach (MySolid mySolid in mySolidsColl)
            {
                // Поиск элементов, которые находятся внутри солидов, построенных на основании пересечения осей
                IEnumerable<Element> intersectedElements = new FilteredElementCollector(_doc, elems.Select(x => x.Id).ToList())
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementIntersectsSolidFilter(mySolid.Solid));

                foreach (Element elem in intersectedElements)
                {
                    if (elem == null) continue;

                    SetParamDataAndDuplicates(elem, _levelParamName, mySolid.LevelIndex);

                    SetParamDataAndDuplicates(elem, _sectionParamName, mySolid.SectionIndex);
                }

                // Получение элементов, которые находятся ВНЕ солида
                notIntersectedElems = notIntersectedElems.Except(intersectedElements, new ElementComparerTool()).ToList();
                PbCounter = elems.Count() - notIntersectedElems.Count();

                pb.Update(PbCounter, "Анализ элементов внутри солидов");
            }

            return notIntersectedElems;
        }

        /// <summary>
        /// Поиск ближайшего солида для элементов, которые не пересекались с солидами
        /// </summary>
        /// <returns>Коллекция элементов, которые не нашли ближайший солид</returns>
        public IEnumerable<Element> FindNearestSolid(
            IEnumerable<Element> notIntersectedElems,
            IEnumerable<MySolid> mySolidsColl,
            Progress_Single pb)

        {
            List<Element> notNearestSolidElems = new List<Element>();

            foreach (Element elem in notIntersectedElems)
            {
                if (elem == null) continue;
                XYZ elemPointCenter = ElemPointCenter(elem);
                if (elemPointCenter == null) continue;

                //Расстояние от центра элемента до центроида солида
                List<MySolid> solidsList = mySolidsColl.ToList();
                solidsList.Sort((x, y) => (int)(x.Solid.ComputeCentroid().DistanceTo(elemPointCenter) - y.Solid.ComputeCentroid().DistanceTo(elemPointCenter)));
                MySolid mySolid = solidsList.FirstOrDefault();

                if (mySolid == null)
                {
                    notNearestSolidElems.Add(elem);
                }

                SetParamDataAndDuplicates(elem, _levelParamName, mySolid.LevelIndex);

                SetParamDataAndDuplicates(elem, _sectionParamName, mySolid.SectionIndex);

                pb.Update(++PbCounter, "Поиск ближайшего солида");
            }

            return notNearestSolidElems;
        }

        /// <summary>
        /// Анализ элементов, которые подверглись двойной записи параметров, и перезапись их
        /// </summary>
        /// <param name="mySolidsColl">Коллекция пользовательских солидов</param>
        /// <param name="pb">Прогресс-бар</param>
        /// <returns>Коллекция элементов, которые не перезаписались</returns>
        public IEnumerable<Element> ReValueDuplicates(
            IEnumerable<MySolid> mySolidsColl,
            Progress_Single pb)
        {

            Dictionary<Element, List<string>> tempDuplicatesWriteParamElems = _duplicatesWriteParamElems
                    //Чистка дубликатов. Больше 4, т.к. минимум вписывается 2 значения - 1 уровень и 2й секция
                    .Where(x => x.Value.Count > 4)
                    .ToDictionary(x => x.Key, x => x.Value);

            _duplicatesWriteParamElems.Clear();

            IEnumerable<Element> notNearestSolidElems = FindNearestSolid(tempDuplicatesWriteParamElems.Keys, mySolidsColl, pb);

            return notNearestSolidElems;
        }

        public void DeleteDirectShapes()
        {
            // Удаляю размещенные солиды
            foreach (DirectShape ds in _directShapesColl)
            {
                _doc.Delete(ds.Id);
            }
        }

        /// <summary>
        /// Поиск номера уровня у элемента (над уровнем) по уровню
        /// </summary>
        private static string GetFloorNumberByLevel(Level lev, int floorTextPosition, char splitChar)
        {
            string levname = lev.Name;
            string[] splitname = levname.Split(splitChar);
            if (splitname.Length < 2)
            {
                throw new Exception($"Некорректное имя уровня: {levname}");
            }
            string floorNumber = splitname[floorTextPosition];

            return floorNumber;
        }

        /// <summary>
        /// Запись данных в параметр, а также наполнение коллекции дубликатов
        /// </summary>
        private void SetParamDataAndDuplicates(Element elem, string paramName, string data)
        {
            Parameter param = elem.LookupParameter(paramName);
            if ((param != null && !param.IsReadOnly))
            {
                param.Set(data);
                if (_duplicatesWriteParamElems.ContainsKey(elem))
                {
                    _duplicatesWriteParamElems[elem].Add(data);
                }
                else
                {
                    _duplicatesWriteParamElems.Add(elem, new List<string>() { data });
                }
            }
        }

        /// <summary>
        /// Точка пересечения осей
        /// </summary>
        /// <param name="grids">Список осей</param>
        /// <returns></returns>
        private static List<XYZ> GetPointsOfGridsIntersection(List<Grid> grids)
        {
            List<XYZ> pointsOfGridsIntersect = new List<XYZ>();
            foreach (Grid grid1 in grids)
            {
                if (grid1 == null) continue;
                Curve curve1 = grid1.Curve;
                foreach (Grid grid2 in grids)
                {
                    if (grid2 == null) continue;
                    if (grid1.Id == grid2.Id) continue;
                    Curve curve2 = grid2.Curve;
                    IntersectionResultArray intersectionResultArray = new IntersectionResultArray();
                    curve1.Intersect(curve2, out intersectionResultArray);

                    if (intersectionResultArray == null || intersectionResultArray.IsEmpty) continue;

                    foreach (IntersectionResult intersection in intersectionResultArray)
                    {
                        XYZ point = intersection.XYZPoint;
                        if (!IsListContainsPoint(pointsOfGridsIntersect, point))
                        {
                            pointsOfGridsIntersect.Add(point);
                        }
                    }
                }
            }
            return pointsOfGridsIntersect;
        }

        private static XYZ ElemPointCenter(Element elem)
        {
            XYZ elemPointCenter = null;

            try
            {
                List<Solid> solidsFromElem = GeometryTool.GetSolidsFromElement(elem);
                int solidsCount = solidsFromElem.Count;
                if (solidsCount > 0)
                {
                    elemPointCenter = solidsCount == 1 ? solidsFromElem.First().ComputeCentroid() : solidsFromElem[solidsCount / 2].ComputeCentroid();
                }
            }
            catch (Exception)
            {
                BoundingBoxXYZ bbox = elem.get_BoundingBox(null);
                elemPointCenter = SolidBoundingBox(bbox).ComputeCentroid();
            }

            return elemPointCenter;
        }

        private static Solid SolidBoundingBox(BoundingBoxXYZ bbox)
        {
            // corners in BBox coords
            XYZ pt0 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
            XYZ pt1 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
            XYZ pt2 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z);
            XYZ pt3 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z);
            //edges in BBox coords
            Line edge0 = Line.CreateBound(pt0, pt1);
            Line edge1 = Line.CreateBound(pt1, pt2);
            Line edge2 = Line.CreateBound(pt2, pt3);
            Line edge3 = Line.CreateBound(pt3, pt0);
            //create loop, still in BBox coords
            List<Curve> edges = new List<Curve>();
            edges.Add(edge0);
            edges.Add(edge1);
            edges.Add(edge2);
            edges.Add(edge3);
            Double height = bbox.Max.Z - bbox.Min.Z;
            CurveLoop baseLoop = CurveLoop.Create(edges);
            List<CurveLoop> loopList = new List<CurveLoop>();
            loopList.Add(baseLoop);
            Solid preTransformBox = GeometryCreationUtilities.CreateExtrusionGeometry(loopList, XYZ.BasisZ, height);

            Solid transformBox = SolidUtils.CreateTransformed(preTransformBox, bbox.Transform);

            return transformBox;
        }

        private static bool IsListContainsPoint(List<XYZ> pointsList, XYZ point)
        {
            foreach (XYZ curpoint in pointsList)
            {
                if (curpoint.IsAlmostEqualTo(point))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Центр между точками
        /// </summary>
        /// <param name="pointsOfGridsIntersect">Список точек</param>
        /// <returns></returns>
        private static XYZ GetCenterPointOfPoints(List<XYZ> pointsOfGridsIntersect)
        {
            double totalX = 0, totalY = 0, totalZ = 0;
            foreach (XYZ xyz in pointsOfGridsIntersect)
            {
                totalX += xyz.X;
                totalY += xyz.Y;
                totalZ += xyz.Z;
            }
            double centerX = totalX / pointsOfGridsIntersect.Count;
            double centerY = totalY / pointsOfGridsIntersect.Count;
            double centerZ = totalZ / pointsOfGridsIntersect.Count;
            return new XYZ(centerX, centerY, centerZ);
        }

        /// <summary>
        /// Создание солида по параметрам
        /// </summary>
        /// <param name="minMaxCoords">Массив из мин и макс точек для солида по оси Z</param>
        /// <param name="pointsOfGridsIntersect">Точки пересечения граничных осей</param>
        /// <param name="solid">Результирующий солид</param>
        /// <returns>Созданная по солиду геометрия</returns>
        /// <exception cref="Exception"></exception>
        private DirectShape CreateSolidsInModel(double[] minMaxCoords, List<XYZ> pointsOfGridsIntersect, out Solid solid)
        {
            List<XYZ> pointsOfGridsIntersectDwn = new List<XYZ>();
            List<XYZ> pointsOfGridsIntersectUp = new List<XYZ>();
            foreach (XYZ point in pointsOfGridsIntersect)
            {
                XYZ newPointDwn = new XYZ(point.X, point.Y, minMaxCoords[0]);
                pointsOfGridsIntersectDwn.Add(newPointDwn);
                XYZ newPointUp = new XYZ(point.X, point.Y, minMaxCoords[1]);
                pointsOfGridsIntersectUp.Add(newPointUp);
            }

            List<Curve> curvesListDwn = GetCurvesListFromPoints(pointsOfGridsIntersectDwn);
            List<Curve> curvesListUp = GetCurvesListFromPoints(pointsOfGridsIntersectUp);
            CurveLoop curveLoopDwn = CurveLoop.Create(curvesListDwn);
            CurveLoop curveLoopUp = CurveLoop.Create(curvesListUp);
            try
            {
                CurveLoop[] curves = new CurveLoop[] { curveLoopDwn, curveLoopUp };
                SolidOptions solidOptions = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                // Отрисовка солидов. Возможно стоит вынести в возможность создания для координаторов???
                //SolidOptions solidOptions = new SolidOptions(new ElementId(6102499), new ElementId(127916));
                solid = GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions);

                DirectShape directShape = DirectShape.CreateElement(_doc, new ElementId(BuiltInCategory.OST_GenericModel));
                directShape.AppendShape(new GeometryObject[] { solid });
                return directShape;
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                throw new Exception("Пограничные оси обязательно должны пересекаться!");
            }
            catch (Exception e)
            {
                PrintError(e);
                solid = null;
                return null;
            }
        }

        private static List<Curve> GetCurvesListFromPoints(List<XYZ> pointsOfGridsIntersect)
        {
            List<Curve> curvesList = new List<Curve>();
            for (int i = 0; i < pointsOfGridsIntersect.Count; i++)
            {
                if (i == pointsOfGridsIntersect.Count - 1)
                {
                    curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[0]));
                    continue;
                }
                curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[i + 1]));

            }
            return curvesList;
        }

        /// <summary>
        /// Минимально нижняя точка и максимально верхняя точка у коллекции элементов
        /// </summary>
        /// <param name="elems">Список элементов</param>
        /// <returns></returns>
        private static double[] GetMinMaxZCoordOfModel(List<Element> elems)
        {
            List<BoundingBoxXYZ> elemsBox = new List<BoundingBoxXYZ>();
            foreach (Element item in elems)
            {
                BoundingBoxXYZ itemBox = item.get_BoundingBox(null);
                if (itemBox == null) continue;
                elemsBox.Add(itemBox);
            }
            double maxPointOfModel = elemsBox.Select(x => x.Max.Z).Max();
            double minPointOfModel = elemsBox.Select(x => x.Min.Z).Min();
            return new double[] { minPointOfModel, maxPointOfModel };
        }

        /// <summary>
        /// Минимально нижняя точка и максимально верхняя точка между уровнями
        /// </summary>
        /// <param name="level">Нижний уровень</param>
        /// <param name="aboveLevel">Верхний уровень</param>
        /// <returns></returns>
        private static double[] GetMinMaxZCoordOfLevel(Level level, Level aboveLevel)
        {
            double minPointOfLevels = level.Elevation;

            double maxPointOLevels;
            if (aboveLevel == null)
            {
                maxPointOLevels = minPointOfLevels + 5000 / 304.8;
            }
            else
            {
                maxPointOLevels = aboveLevel.Elevation;
            }

            return new double[] { minPointOfLevels, maxPointOLevels };
        }
    }
}
