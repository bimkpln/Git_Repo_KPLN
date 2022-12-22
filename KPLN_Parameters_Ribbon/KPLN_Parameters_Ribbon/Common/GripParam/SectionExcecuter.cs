﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using KPLN_Parameters_Ribbon.Common.Tools;
using static KPLN_Loader.Output.Output;
using Autodesk.Revit.UI;
using KPLN_Parameters_Ribbon.Forms;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Общий класс для обработки секций
    /// </summary>
    internal static class SectionExcecuter
    {
        /// <summary>
        /// Общие метод записи значения секции для всех элементов
        /// </summary>
        /// <param name="doc">Документ Ревит</param>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="gridSectionParam">Имя параметра номера секции у осей</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <returns></returns>
        public static bool ExecuteByElement(Document doc, List<Element> elems, string gridSectionParam, string sectionParam, Progress_Single pb)
        {
            double[] minMaxCoords = GetMinMaxZCoordOfModel(elems);
            List<Grid> grids = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Grid)).Cast<Grid>().ToList();
            
            Dictionary<string, HashSet<Grid>> sectionsGrids = new Dictionary<string, HashSet<Grid>>();
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
            if (sectionsGrids.Keys.Count == 0)
            {
                throw new Exception($"Для заполнения номера секции в элементах, необходимо заполнить параметр: {gridSectionParam} в осях! Значение указывается через \"-\" для осей, относящихся к нескольким секциям.");
            }

            #region Этап №1 - анализ на коллизию с солидом
            List<Element> notIntersectedElems = elems;
            Dictionary<Solid, string> sectionSolids = new Dictionary<Solid, string>();
            Dictionary<Element, List<string>> duplicatesWriteElems = new Dictionary<Element, List<string>>(new ElementComparerTool());
            List<DirectShape> directShapes = new List<DirectShape>();
            foreach (string sg in sectionsGrids.Keys)
            {
                List<Grid> gridsOfSect = sectionsGrids[sg].ToList();
                if (sectionsGrids[sg].Count < 4)
                {
                    throw new Exception($"Количество осей с номером секции: {sg} меньше 4. Проверьте назначение параметров у осей!");
                }
                List<XYZ> pointsOfGridsIntersect = GetPointsOfGridsIntersection(gridsOfSect);
                pointsOfGridsIntersect.Sort(new ClockwiseComparerTool(GetCenterPointOfPoints(pointsOfGridsIntersect)));
                
                Solid solid;
                DirectShape directShape = CreateSolidsInModel(doc, minMaxCoords, pointsOfGridsIntersect, out solid);
                sectionSolids.Add(solid, sg);
                directShapes.Add(directShape);
                
                // Поиск элементов, которые находятся внутри солидов, построенных на основании пересечения осей
                List<Element> intersectedElements = new FilteredElementCollector(doc, elems.Select(x => x.Id).ToList())
                    .WhereElementIsNotElementType()
                    .WherePasses(new ElementIntersectsSolidFilter(solid))
                    .ToElements()
                    .ToList();
                foreach (Element item in intersectedElements)
                {
                    if (item == null) continue;
                    
                    Parameter parameter = item.LookupParameter(sectionParam);
                    if (parameter != null && !parameter.IsReadOnly)
                    {
                        parameter.Set(sg);
                        if (duplicatesWriteElems.ContainsKey(item))
                        {
                            duplicatesWriteElems[item].Add(sg);
                        }
                        else
                        {
                            duplicatesWriteElems.Add(item, new List<string>() { sg });
                            
                            pb.Increment();
                        }
                    }
                }
                
                // Получение элементов, которые находятся ВНЕ солида
                notIntersectedElems = notIntersectedElems.Except(intersectedElements, new ElementComparerTool()).ToList();
            }
            
            duplicatesWriteElems = duplicatesWriteElems.Where(x => x.Value.Count > 1).ToDictionary(x => x.Key, x => x.Value);
            Print($"Количество элементов, которые подверглись перезаписи параметра на этапе №1 (поиск внутри пересечения осей): {duplicatesWriteElems.Keys.Count}." +
                $"\nОни подвеграются вторичному анализу", KPLN_Loader.Preferences.MessageType.Regular);
            // Осуществляю поиск ближайшего ОДНОГО солида для элементов с двойной записью парамтеров
            foreach (KeyValuePair<Element, List<string>> item in duplicatesWriteElems)
            {
                if (item.Key == null) continue;
                XYZ elemPointCenter = ElemPointCenter(item.Key);
                if (elemPointCenter == null) continue;

                //Расстояние от центра элемента до центроида солида
                List<Solid> solidsList = sectionSolids.Where(x => item.Value.Contains(x.Value)).ToDictionary(x => x.Key, x => x.Value).Keys.ToList();
                solidsList.Sort((x, y) => (int)(x.ComputeCentroid().DistanceTo(elemPointCenter) - y.ComputeCentroid().DistanceTo(elemPointCenter)));
                Solid solid = solidsList.First();
                if (solid == null) continue;
                
                Parameter parameter = item.Key.LookupParameter(sectionParam);
                if (parameter != null && !parameter.IsReadOnly)
                {
                    parameter.Set(sectionSolids[solid]);

                    pb.Increment();
                }
            }

            Print($"Количество необработанных элементов после 1-ого этапа (поиск внутри пересечения осей): {notIntersectedElems.Count}", KPLN_Loader.Preferences.MessageType.Warning);
            if (notIntersectedElems.Count == 0)
            {
                return true;
            }
            #endregion

            #region Этап №2 - анализ остатка
            pb.Decrement(notIntersectedElems.Count);
            Print($"\nОсуществляю поиск ближайшей секции\n", KPLN_Loader.Preferences.MessageType.Regular);
            List<Element> notNearestSolidElems = notIntersectedElems;
            foreach (Element elem in notIntersectedElems)
            {
                if (elem == null) continue;
                XYZ elemPointCenter = ElemPointCenter(elem);
                if (elemPointCenter == null) continue;

                //Расстояние от центра элемента до центроида солида
                List<Solid> solidsList = sectionSolids.Keys.ToList();
                solidsList.Sort((x, y) => (int)(x.ComputeCentroid().DistanceTo(elemPointCenter) - y.ComputeCentroid().DistanceTo(elemPointCenter)));
                Solid solid = solidsList.First();
                if (solid == null) continue;
                
                Parameter parameter = elem.LookupParameter(sectionParam);
                if (parameter != null && !parameter.IsReadOnly)
                {
                    if (parameter.Set(sectionSolids[solid]))
                    {
                        // Получение элементов, которые не могут найти ближайший элемент
                        notNearestSolidElems = notNearestSolidElems.Except(notIntersectedElems, new ElementComparerTool()).ToList();

                        pb.Increment();
                    }
                }
            }
            if(notNearestSolidElems.Count > 0)
            {
                Print($"Количество необработанных элементов после 2-ого этапа (поиск ближайшей секции): {notNearestSolidElems.Count}", KPLN_Loader.Preferences.MessageType.Warning);
                foreach(Element element in notNearestSolidElems)
                {
                    Print($"Проверь вручную элемент с id: {element.Id}", KPLN_Loader.Preferences.MessageType.Warning);
                }

            }
            #endregion

            // Удаляю размещенные солиды
            foreach (DirectShape ds in directShapes)
            {
                doc.Delete(ds.Id);
            }
            return true;
        }

        /// <summary>
        /// Общие метод записи значения секции для витражей АР
        /// </summary>
        /// <param name="doc">Документ Ревит</param>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <returns></returns>
        public static bool ExecuteByHost_AR(List<Element> elems, string sectionParam, Progress_Single pb)
        {
            foreach(Element elem in elems)
            {
                Element hostElem = null;
                Wall hostWall = null;
                Type elemType = elem.GetType();
                switch (elemType.Name)
                {
                    // Проброс панелей витража на стену
                    case nameof(Panel):
                        Panel panel = (Panel)elem;
                        hostWall = (Wall)panel.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс импостов витража на стену
                    case nameof(Mullion):
                        Mullion mullion = (Mullion)elem;
                        hostWall = (Wall)mullion.Host;
                        hostElem = hostWall;
                        break;

                    // Проброс на вложенные общие семейства
                    default:
                        FamilyInstance instance = elem as FamilyInstance;
                        hostElem = instance.SuperComponent;
                        break;
                }
                
                var hostElemParamValue = hostElem.LookupParameter(sectionParam).AsString();
                elem.LookupParameter(sectionParam).Set(hostElemParamValue);

                pb.Increment();
            }
            return true;
        }

        /// <summary>
        /// Общие метод записи значения секции для вложенных общих семейств
        /// </summary>
        /// <param name="doc">Документ Ревит</param>
        /// <param name="elems">Коллекция элементов для обработки</param>
        /// <param name="sectionParam">Имя параметра номера секции у заполняемых элементов</param>
        /// <returns></returns>
        public static bool ExecuteByHost(List<Element> elems, string sectionParam, Progress_Single pb)
        {
            foreach (Element elem in elems)
            {
                FamilyInstance instance = elem as FamilyInstance;
                Element hostElem = instance.SuperComponent;
                var hostElemParamValue = hostElem.LookupParameter(sectionParam).AsString();
                elem.LookupParameter(sectionParam).Set(hostElemParamValue);

                pb.Increment();
            }
            return true;
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

        private static DirectShape CreateSolidsInModel(Document doc, double[] minMaxCoords, List<XYZ> pointsOfGridsIntersect, out Solid solid)
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
                solid = GeometryCreationUtilities.CreateLoftGeometry(curves, solidOptions);
                DirectShape directShape = DirectShape.CreateElement(doc, new ElementId(BuiltInCategory.OST_GenericModel));
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
                    try
                    {
                        curve1.Intersect(curve2, out intersectionResultArray);
                    }
                    catch (Exception e)
                    {
                        TaskDialog.Show("Отладка", e.ToString());
                    }

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
                try
                {
                    if (i == pointsOfGridsIntersect.Count - 1)
                    {
                        curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[0]));
                        continue;
                    }
                    curvesList.Add(Line.CreateBound(pointsOfGridsIntersect[i], pointsOfGridsIntersect[i + 1]));
                }
                catch (Exception ex) { PrintError(ex); }
            }
            return curvesList;
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

        public static Solid SolidBoundingBox(BoundingBoxXYZ bbox)
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

    }
}
