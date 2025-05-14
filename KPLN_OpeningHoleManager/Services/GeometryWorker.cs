using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class GeometryWorker
    {
        /// <summary>
        /// Создать СОЛИД вручную. Выдавливание идёт вдоль оси Z
        /// </summary>
        /// <returns></returns>
        internal static Solid CreateSolid_ZDir(XYZ direction, XYZ insertPoint, double height, double width, double radius)
        {
            double extrusionDist = height > 0 ? height : radius; 

            // Сістэма каардынат
            XYZ xDir = direction; // па шырыні
            XYZ zDir = XYZ.BasisZ; // уверх
            XYZ yDir = zDir.CrossProduct(xDir).Normalize(); // бок

            List<Curve> curvesList;
            if (radius == 0)
            {
                List<XYZ> points = new List<XYZ>
                {
                    insertPoint - xDir * width/2,
                    insertPoint - xDir * width/2 + yDir * height,
                    insertPoint + xDir * width/2 + yDir * height,
                    insertPoint + xDir * width/2,
                };

                curvesList = GetCurvesList_LinesByPoints(points);
            }
            else
                curvesList = GetCurvesList_ArcByInsertPntAndRadius(direction, insertPoint, radius);

            CurveLoop[] curves = new CurveLoop[] { CurveLoop.Create(curvesList) };
            Solid extrSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curves, zDir, extrusionDist);

            //var centr1 = extrSolid.ComputeCentroid();
            //var bbox1 = extrSolid.GetBoundingBox();
            //var trans1 = bbox1.Transform;
            //var newbbox1 = new BoundingBoxXYZ() { Min = trans1.OfPoint(bbox1.Min), Max = trans1.OfPoint(bbox1.Max) };

            return extrSolid;
        }

        /// <summary>
        /// Архвная копия метода - выдавливание по перпендикулярной оси для вектора direction
        /// </summary>
        /// <returns></returns>
        internal static Solid CreateSolid_XYDir(XYZ direction, XYZ insertPoint, double height, double width, double radius)
        {
            // Нармалізуем напрамак
            XYZ xDir = direction.Normalize(); // напрамак шырыні
            XYZ upGuess = XYZ.BasisZ;
            if (Math.Abs(xDir.DotProduct(upGuess)) > 0.99)
                upGuess = XYZ.BasisX;

            XYZ yDir = upGuess.CrossProduct(xDir).Normalize(); // напрамак экструзіі (перпендыкуляр)
            XYZ zDir = xDir.CrossProduct(yDir).Normalize();    // уверх

            List<Curve> curvesList;
            if (radius == 0)
            {
                List<XYZ> points = new List<XYZ>
                {
                    insertPoint - xDir * width/2 - zDir * height/2,
                    insertPoint - xDir * width/2 + zDir * height/2,
                    insertPoint + xDir * width/2 + zDir * height/2,
                    insertPoint + xDir * width/2 - zDir * height/2,
                };

                curvesList = GetCurvesList_LinesByPoints(points);
            }
            else
                curvesList = GetCurvesList_ArcByInsertPntAndRadius(zDir, insertPoint, radius);

            CurveLoop[] curves = new CurveLoop[] { CurveLoop.Create(curvesList) };
            Solid extrSolid = GeometryCreationUtilities.CreateExtrusionGeometry(curves, yDir, 1);

            return extrSolid;
        }


        /// <summary>
        /// Получить SOLID элемента Ревит
        /// </summary>
        internal static Solid GetRevitElemSolid(Element elem, Transform transform = null)
        {
            Solid resultSolid = null;

            // Для кабельных лотков лучше подходит средний ур. детализации
            Options opt = new Options();
            if (elem is CableTray)
                opt.DetailLevel = ViewDetailLevel.Medium;
            else
                opt.DetailLevel = ViewDetailLevel.Fine;

            GeometryElement geomElem = elem.get_Geometry(opt);
            foreach (GeometryObject gObj in geomElem)
            {
                Solid solid = gObj as Solid;
                GeometryInstance gInst = gObj as GeometryInstance;
                if (solid != null)
                    resultSolid = solid;
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
                            resultSolid = solid;
                        }
                    }
                }
            }

            if (resultSolid == null && (elem is FamilyInstance fi && fi.SuperComponent == null))
                throw new Exception($"У элемента с id: {elem.Id} из модели {elem.Document.Title} проблемы с получением Solid. Отправь разработчику.");

            // Трансформ по координатам (если нужно)
            if (resultSolid != null && transform != null)
                resultSolid = SolidUtils.CreateTransformed(resultSolid, transform);

            return resultSolid;
        }

        public static double GetMinimumDistanceBetweenSolids(Solid solid1, Solid solid2)
        {
            double minDistance = double.MaxValue;

            // Прайдзіся па ўсіх гранях першага соліда
            foreach (Face face1 in solid1.Faces)
            {
                Mesh mesh1 = face1.Triangulate();

                foreach (XYZ v1 in mesh1.Vertices)
                {
                    // Для кожнай кропкі на першым солідзе, вызнач бліжэйшую кропку на другім
                    foreach (Face face2 in solid2.Faces)
                    {
                        IntersectionResult intResult = face2.Project(v1);
                        if (intResult != null)
                        {
                            // Праектуем кропку на паверхню другога соліда
                            XYZ projected = intResult.XYZPoint;
                            double distance = v1.DistanceTo(projected);

                            if (distance < minDistance)
                                minDistance = distance;
                        }
                    }
                }
            }

            return minDistance;
        }

        /// <summary>
        /// Получить общий BoundingBoxXYZ на основе всех заданий от ИОС с УЧЁТОМ вшитого Transform
        /// </summary>
        /// <returns></returns>
        internal static BoundingBoxXYZ CreateOverallBBox(IOSOpeningHoleTaskEntity[] iosTasks)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (IOSOpeningHoleTaskEntity iosTask in iosTasks)
            {
                BoundingBoxXYZ elementBox = iosTask.OHE_Element.get_BoundingBox(null);
                if (elementBox == null)
                {
                    HtmlOutput.Print($"Ошибка анализа, могут быть не предвиденные результат (проверь отдельно)." +
                        $" У элемента с id:{iosTask.OHE_Element.Id} из связи {iosTask.OHE_LinkDocument.Title} - нет BoundingBoxXYZ",
                        MessageType.Error);
                    continue;
                }

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = iosTask.OHE_LinkTransform.OfPoint(elementBox.Min),
                        Max = iosTask.OHE_LinkTransform.OfPoint(elementBox.Max)
                    };
                }
                else
                {
                    resultBBox.Min = iosTask.OHE_LinkTransform.OfPoint(new XYZ(
                        Math.Min(resultBBox.Min.X, elementBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elementBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elementBox.Min.Z)));

                    resultBBox.Max = iosTask.OHE_LinkTransform.OfPoint(new XYZ(
                        Math.Max(resultBBox.Max.X, elementBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elementBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elementBox.Max.Z)));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Создать Outline по указанному BoundingBoxXYZ и расширению
        /// </summary>
        internal static Outline CreateFilterOutline(BoundingBoxXYZ bbox, double expandValue)
        {
            Outline resultOutlie;

            XYZ bboxMin = bbox.Min;
            XYZ bboxMax = bbox.Max;

            // Подготовка расширенного BoundingBoxXYZ, чтобы не упустить эл-ты
            BoundingBoxXYZ expandedCropBB = new BoundingBoxXYZ()
            {
                Max = bboxMax + new XYZ(expandValue, expandValue, expandValue),
                Min = bboxMin - new XYZ(expandValue, expandValue, expandValue),
            };

            resultOutlie = new Outline(expandedCropBB.Min, expandedCropBB.Max);

            if (resultOutlie.IsEmpty)
            {
                XYZ transExpandedElemBBMin = resultOutlie.MinimumPoint;
                XYZ transExpandedElemBBMax = resultOutlie.MaximumPoint;

                double minX = transExpandedElemBBMin.X;
                double minY = transExpandedElemBBMin.Y;
                double minZ = transExpandedElemBBMin.Z;

                double maxX = transExpandedElemBBMax.X;
                double maxY = transExpandedElemBBMax.Y;
                double maxZ = transExpandedElemBBMax.Z;

                double sminX = Math.Min(minX, maxX);
                double sminY = Math.Min(minY, maxY);
                double sminZ = Math.Min(minZ, maxZ);

                double smaxX = Math.Max(minX, maxX);
                double smaxY = Math.Max(minY, maxY);
                double smaxZ = Math.Max(minZ, maxZ);

                XYZ pntMin = new XYZ(sminX, sminY, sminZ);
                XYZ pntMax = new XYZ(smaxX, smaxY, smaxZ);

                resultOutlie = new Outline(pntMin, pntMax);
            }

            if (!resultOutlie.IsEmpty && resultOutlie.IsValidObject)
                return resultOutlie;

            HtmlOutput.Print($"Отправь разработчику - не удалось создать Outline для фильтрации", MessageType.Error);
            return null;
        }

        /// <summary>
        /// Получить вектор для основы отверстия
        /// </summary>
        internal static XYZ GetHostDirection(Element host)
        {
            // Получаю вектор для стены
            XYZ wallDirection = null;
            if (host is Wall wall)
            {
                Curve curve = (wall.Location as LocationCurve).Curve ??
                    throw new Exception($"Не обработанная основа (не Curve) с id: {host.Id}. Отправь разработчику!");

                LocationCurve wallLocCurve = wall.Location as LocationCurve;
                if (wallLocCurve.Curve is Line wallLine)
                    wallDirection = wallLine.Direction;
                else
                    throw new Exception($"Пока умею работать только с прямыми стенами :(");


            }
            else
                throw new Exception($"Не обработанная основа с id: {host.Id}. Отправь разработчику!");


            return wallDirection;
        }


        /// <summary>
        /// Получить плоскость у солида с определенным углом к указанному вектору 
        /// </summary>
        internal static Face GetFace_ByAngleToDirection(Solid checkSolid, XYZ hostDirection, double checkAngle = 90)
        {
            Face result = null;

            double tempArea = 0;
            foreach (Face face in checkSolid.Faces)
            {
                if (face is PlanarFace planarFace)
                {
                    XYZ faceNormal = planarFace.FaceNormal;
                    XYZ checkOrigin = new XYZ(faceNormal.X, faceNormal.Y, hostDirection.Z);

#if Debug2020 || Revit2020
                    double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), DisplayUnitType.DUT_DEGREES_AND_MINUTES);
#else
                    double angle = UnitUtils.ConvertFromInternalUnits(hostDirection.AngleTo(checkOrigin), new ForgeTypeId("autodesk.unit.unit:degrees-1.0.1"));
#endif
                    if (Math.Round(angle, 5) == checkAngle && tempArea < face.Area)
                    {
                        result = face;
                        tempArea = face.Area;
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// Получить высоту и ширину (именно в таком порядке) для СОЛИДА относительно вектора размещения солида
        /// </summary>
        /// <param name="solid">Солид для анализа</param>
        /// <param name="mainDirection">Вектор, на который проецируем точки</param>
        /// <returns></returns>
        internal static double[] GetSolidWidhtAndHeight_ByDirection(Solid solid, XYZ mainDirection)
        {
            // Атрымліваем дзве артагональныя восі
            XYZ up = XYZ.BasisZ;
            if (Math.Abs(mainDirection.DotProduct(up)) > 0.99)
                up = XYZ.BasisX; // калі mainDirection амаль Z, то выбіраем іншы
            XYZ sideDirection = mainDirection.CrossProduct(up);
            up = sideDirection.CrossProduct(mainDirection);

            // Збіраем усе вяршыні
            List<XYZ> allPoints = new List<XYZ>();
            foreach (Face face in solid.Faces)
            {
                foreach (EdgeArray edgeArray in face.EdgeLoops)
                {
                    foreach (Edge edge in edgeArray)
                    {
                        allPoints.AddRange(edge.Tessellate());
                    }
                }
            }

            // Праектуем усе пункты ў мясцовую сістэму каардынат
            List<XYZ> projected = allPoints
                .Select(p =>
                    {
                        double x = p.DotProduct(mainDirection);
                        double y = p.DotProduct(sideDirection);
                        double z = p.DotProduct(up);
                        return new XYZ(x, y, z);
                    })
                .ToList();

            // Знаходзім памеры
            double minX = projected.Min(p => p.X);
            double maxX = projected.Max(p => p.X);
            double minY = projected.Min(p => p.Y);
            double maxY = projected.Max(p => p.Y);
            double minZ = projected.Min(p => p.Z);
            double maxZ = projected.Max(p => p.Z);

            double width = maxX - minX;
            double height = maxZ - minZ;

            double[] result = new double[2]
            {
                width,
                height
            };

            return result;
        }

        /// <summary>
        /// Получить коллекцию кривых из линий по точкам
        /// </summary>
        private static List<Curve> GetCurvesList_LinesByPoints(List<XYZ> pointsOfIntersect)
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
        /// Получить коллекцию кривых из дуги по точкам
        /// </summary>
        private static List<Curve> GetCurvesList_ArcByInsertPntAndRadius(XYZ direction, XYZ insertPoint, double radius)
        {
            // Нармалізуем напрамак
            XYZ xDir = direction.Normalize(); // асноўны напрамак (дыяметр)

            // Знаходзім вектар, перпендыкулярны xDir у плоскасці (напрыклад, XY)
            XYZ upGuess = XYZ.BasisZ;
            if (Math.Abs(xDir.DotProduct(upGuess)) > 0.99)
                upGuess = XYZ.BasisX;

            XYZ yDir = upGuess.CrossProduct(xDir).Normalize(); // перпендыкуляр у плоскасці

            // Пачатковая і канчатковая кропка круга
            XYZ ptStart = insertPoint - xDir * radius;
            XYZ ptEnd = insertPoint + xDir * radius;

            // Дапаможныя сярэднія кропкі, каб задаць паўкругі
            XYZ ptMid = insertPoint + yDir * radius;
            XYZ ptMid2 = insertPoint - yDir * radius;

            Arc arc1 = Arc.Create(ptStart, ptEnd, ptMid);
            Arc arc2 = Arc.Create(ptEnd, ptStart, ptMid2);

            return new List<Curve> { arc1, arc2 };
        }
    }
}
