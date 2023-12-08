using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil.Common;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KPLN_ModelChecker_Lib
{
    /// <summary>
    /// Класс для генерации солида между уровнем и оргаждающими секциями
    /// </summary>
    public class LevelAndGridSolid
    {
        /// <summary>
        /// Солид уровня
        /// </summary>
        public Solid LevelSolid { get; private set; }

        /// <summary>
        /// Ссылка на текущий CheckLevelOfInstanceLevelData
        /// </summary>
        public LevelData CurrentLevelData { get; private set; }

        /// <summary>
        /// Ссылка на CheckLevelOfInstanceGridData
        /// </summary>
        public GridData GridData { get; private set; }

        private LevelAndGridSolid(Solid solid, LevelData currentLevel, GridData gData)
        {
            LevelSolid = solid;
            CurrentLevelData = currentLevel;
            GridData = gData;
        }

        /// <summary>
        /// Подготовка коллекции солидов с данными по секциям
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="gridSeparParamName">Параметр для разделения осей по секциям</param>
        public static List<LevelAndGridSolid> PrepareSolids(Document doc, string gridSeparParamName)
        {
            List<LevelAndGridSolid> result = new List<LevelAndGridSolid>();

            List<GridData> gridDatas = GridData.GridPrepare(doc, gridSeparParamName);
            List<LevelData> levelDatas = LevelData.LevelPrepare(doc);

            // Подготовка предварительной коллекции элементов
            List<LevelAndGridSolid> preResult = new List<LevelAndGridSolid>();
            foreach (LevelData currentLevel in levelDatas)
            {
                bool isCreated = false;
                foreach (GridData gData in gridDatas)
                {
                    if (currentLevel.CurrentSectionNumber.Equals(gData.CurrentSection))
                    {
                        Solid levSolid = CreateSolidInModel(currentLevel, gData);
                        LevelAndGridSolid secData = new LevelAndGridSolid(levSolid, currentLevel, gData);
                        preResult.Add(secData);
                        isCreated = true;
                    }
                }

                if (!isCreated)
                    throw new CheckerException($"Проблема с несовпадением названий секций в уровнях и в осях - нужно синхронизировать данные");
            }

            // Очистка от солидов, для вспомогательных уровней внутри секций
            foreach (LevelAndGridSolid secData in preResult)
            {
                IEnumerable<LevelAndGridSolid> currentSectionAndAboveLevelsColl = null;
                if (secData.CurrentLevelData.CurrentAboveLevel != null)
                {
                    currentSectionAndAboveLevelsColl = preResult
                        .Where(r =>
                            r.GridData.CurrentSection.Equals(secData.GridData.CurrentSection)
                            && r.CurrentLevelData.CurrentAboveLevel != null
                            && r.CurrentLevelData.CurrentAboveLevel.Id == secData.CurrentLevelData.CurrentAboveLevel.Id);
                }
                else
                {
                    currentSectionAndAboveLevelsColl = preResult
                        .Where(r =>
                            r.GridData.CurrentSection.Equals(secData.GridData.CurrentSection)
                            && r.CurrentLevelData.CurrentAboveLevel == null);
                }

                if (currentSectionAndAboveLevelsColl.Count() == 1)
                    result.Add(secData);
                else if (currentSectionAndAboveLevelsColl.Any())
                {
                    LevelAndGridSolid minSecData = currentSectionAndAboveLevelsColl
                        .Aggregate((lvlMinElv, x) => (x.CurrentLevelData.CurrentLevel.Elevation < lvlMinElv.CurrentLevelData.CurrentLevel.Elevation) ? x : lvlMinElv);
                    if (!result.Contains(minSecData))
                        result.Add(minSecData);
                }
                else
                {
                    result.Add(secData);
                }
            }

            // Проверка на коллизии между Solid
            foreach (LevelAndGridSolid secData1 in result)
            {
                foreach (LevelAndGridSolid secData2 in result)
                {
                    if (!secData1.Equals(secData2))
                    {
                        Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(secData1.LevelSolid, secData2.LevelSolid, BooleanOperationsType.Intersect);
                        if (intersectionSolid != null && intersectionSolid.Volume > 0)
                            throw new CheckerException("Солиды уровней пересекаются (ошибка в заполнении параметров сепарации объекта, либо уровни названы не по BEP). Отправь разработчику: " +
                                $"Уровень id: {secData1.CurrentLevelData.CurrentLevel.Id} и {secData2.CurrentLevelData.CurrentLevel.Id} " +
                                $"для секции №{secData1.GridData.CurrentSection}");
                    }
                }
            }

            return result;

        }

        /// <summary>
        /// Создание Solid по параметрам
        /// </summary>
        /// <param name="levData">Данные по уровням</param>
        /// <param name="grData">Данные по осям</param>
        /// <returns>Созданная по солиду геометрия</returns>
        private static Solid CreateSolidInModel(LevelData levData, GridData grData)
        {
            List<XYZ> pointsOfGridsIntersect = GetPointsOfGridsIntersection(grData.CurrentGrids);
            pointsOfGridsIntersect.Sort(new PntComparer(GetCenterPointOfPoints(pointsOfGridsIntersect)));

            List<XYZ> pointsOfGridsIntersectDwn = new List<XYZ>();
            List<XYZ> pointsOfGridsIntersectUp = new List<XYZ>();
            foreach (XYZ point in pointsOfGridsIntersect)
            {
                XYZ newPointDwn = new XYZ(point.X, point.Y, levData.MinAndMaxLvlPnts[0]);
                pointsOfGridsIntersectDwn.Add(newPointDwn);
                XYZ newPointUp = new XYZ(point.X, point.Y, levData.MinAndMaxLvlPnts[1]);
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
                return GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                StringBuilder sb = new StringBuilder();
                foreach (Grid gr in grData.CurrentGrids)
                {
                    sb.Append(gr.Id.ToString());
                    sb.Append(", ");
                }

                throw new CheckerException($"Пограничные оси обязательно должны пересекаться! Проверь оси id: {sb.ToString().TrimEnd(", ".ToArray())}");
            }
        }

        /// <summary>
        /// Точка пересечения осей
        /// </summary>
        /// <param name="grids">Список осей</param>
        private static List<XYZ> GetPointsOfGridsIntersection(HashSet<Grid> grids)
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
                        if (!pointsOfGridsIntersect.Any(pgi => point.IsAlmostEqualTo(pgi)))
                            pointsOfGridsIntersect.Add(point);
                    }
                }
            }
            return pointsOfGridsIntersect;
        }

        /// <summary>
        /// Центр между точками
        /// </summary>
        /// <param name="pointsOfGridsIntersect">Список точек</param>
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
    }
}
