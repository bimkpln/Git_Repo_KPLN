using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace KPLN_Looker.Services
{
    /// <summary>
    /// Сервис по генерации списка пересекаемых элементов
    /// </summary>
    internal sealed class IntersectedElemFinderService
    {
        private static int[] _builtInCatIDs;

        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        private static readonly List<BuiltInCategory> _builtInCategories = new List<BuiltInCategory>()
        { 
            // ОВВК (ЭОМСС - огнезащита)
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_PipeCurves,
            // ЭОМСС
            BuiltInCategory.OST_CableTray,
        };

        /// <summary>
        /// Список id BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        public static int[] BuiltInCatIDs 
        {
            get
            {
                if (_builtInCatIDs == null)
                    _builtInCatIDs = _builtInCategories.Select(bic => (int)bic).ToArray();

                return _builtInCatIDs;
            }
        }

        /// <summary>
        /// Подготовка коллекции точек пересечений с добавленными эл-тами
        /// </summary>
        public static List<XYZ> GetIntersectPoints(List<Element> addedLinearElems, Document checkDoc)
        {
            List<XYZ> intersectedPoints = new List<XYZ>();

            Transform addedDocTrans = addedLinearElems.FirstOrDefault().Document.ActiveProjectLocation.GetTotalTransform();
            Transform checkTrans = checkDoc.ActiveProjectLocation.GetTotalTransform();
            // Ищу результирующий Transform. Inverse родной, чтобы выяснить разницу
            Transform resultTransform = addedDocTrans.Inverse * checkTrans;

            foreach (Element addedElem in addedLinearElems)
            {
                BoundingBoxXYZ addedElemBB = addedElem.get_BoundingBox(null);
                if (addedElemBB == null)
                {
                    HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением BoundingBoxXYZ. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                    continue;
                }

                Solid addedElemSolid = GetElemSolid(addedElem);
                if (addedElemSolid == null)
                {
                    HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                    continue;
                }

                Element[] potentialIntersectElems = GetPotentioalIntersectedElems(checkDoc, resultTransform, addedElemBB);
                foreach(Element elemToCheck in potentialIntersectElems)
                {
                    if (elemToCheck.Id.IntegerValue == addedElem.Id.IntegerValue)
                        continue;

                    Solid solidToCheck = GetElemSolid(elemToCheck, resultTransform);
                    Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(addedElemSolid, solidToCheck, BooleanOperationsType.Intersect);
                    if (intersectionSolid != null && intersectionSolid.Volume > 0)
                    {
                        intersectedPoints.Add(intersectionSolid.ComputeCentroid());
                    }
                }
            }

            return intersectedPoints;
        }

        /// <summary>
        /// Подготовить коллекцию ПОТЕНЦИАЛЬНО пересекаемых элементов
        /// </summary>
        /// <param name="checkDoc"></param>
        /// <param name="checkDocTransform"></param>
        /// <param name="addedElemBB"></param>
        /// <returns></returns>
        private static Element[] GetPotentioalIntersectedElems(Document checkDoc, Transform checkDocTransform, BoundingBoxXYZ addedElemBB)
        {
            List<Element> potentialIntersectedElems = new List<Element>();

            // Подготовка расширенного BoundingBoxXYZ, чтобы не упустить эл-ты
            double expandValue = 1;
            XYZ addedElemBBMaxPnt = addedElemBB.Max;
            XYZ addedElemBBMinPnt = addedElemBB.Min;
            BoundingBoxXYZ expandedElemBB = new BoundingBoxXYZ()
            {
                Max = new XYZ(
                    addedElemBBMaxPnt.X < 0 ? addedElemBBMaxPnt.X - expandValue : addedElemBBMaxPnt.X + expandValue,
                    addedElemBBMaxPnt.Y < 0 ? addedElemBBMaxPnt.Y - expandValue : addedElemBBMaxPnt.Y + expandValue,
                    addedElemBBMaxPnt.Z + expandValue),
                Min = new XYZ(
                    addedElemBBMinPnt.X < 0 ? addedElemBBMinPnt.X - expandValue : addedElemBBMinPnt.X + expandValue,
                    addedElemBBMinPnt.Y < 0 ? addedElemBBMinPnt.Y - expandValue : addedElemBBMinPnt.Y + expandValue,
                    addedElemBBMinPnt.Z - expandValue),
            };

            Outline filterOutline = new Outline(
                    checkDocTransform.OfPoint(expandedElemBB.Min),
                    checkDocTransform.OfPoint(expandedElemBB.Max));
            // Проверяю Outline. Если пустой - значит XY нужно флипнуть
            if (filterOutline.IsEmpty)
            {
                XYZ transExpandedElemBBMin = checkDocTransform.OfPoint(expandedElemBB.Min);
                XYZ transExpandedElemBBMax = checkDocTransform.OfPoint(expandedElemBB.Max);

                double minX = transExpandedElemBBMin.X;
                double minY = transExpandedElemBBMin.Y;

                double maxX = transExpandedElemBBMax.X;
                double maxY = transExpandedElemBBMax.Y;

                double sminX = Math.Min(minX, maxX);
                double sminY = Math.Min(minY, maxY);

                double smaxX = Math.Max(minX, maxX);
                double smaxY = Math.Max(minY, maxY);

                XYZ pntMin = new XYZ(sminX, sminY, transExpandedElemBBMin.Z);
                XYZ pntMax = new XYZ(smaxX, smaxY, transExpandedElemBBMax.Z);

                filterOutline = new Outline(pntMin, pntMax);
            }

            // Подготовка коллекции эл-в пересекаемых и внутри расширенного BoundingBoxXYZ
            foreach (BuiltInCategory category in _builtInCategories)
            {
                FilteredElementCollector catFIC = new FilteredElementCollector(checkDoc)
                    .OfCategory(category)
                    .WhereElementIsNotElementType();

                potentialIntersectedElems.AddRange(catFIC
                    .WherePasses(new BoundingBoxIntersectsFilter(filterOutline, 0.1)));

                potentialIntersectedElems.AddRange(catFIC
                    .WherePasses(new BoundingBoxIsInsideFilter(filterOutline, 0.1)));
            }

            return potentialIntersectedElems.Distinct().ToArray();
        }

        /// <summary>
        /// Получить Solid элемента
        /// </summary>
        public static Solid GetElemSolid(Element elem, Transform transform = null)
        {
            Solid resultSolid = null;
            GeometryElement geomElem = elem.get_Geometry(new Options { DetailLevel = ViewDetailLevel.Fine });
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

            // ВАЖНО: Солиды у системных линейных элементов должны быть всегда. Если проверка расширеться на FamilyInstance - нужно будет фильтровать эл-ты без геометрии
            if (resultSolid == null)
            {
                HtmlOutput.Print($"У элемента с id: {elem.Id} из модели {elem.Document.Title} проблемы с получением Solid. " +
                    $"Элемент проигнорирован, но ошибку стоит отправить разработчику);", MessageType.Error);

                return null;
            }

            // Трансформ по координатам (если нужно)
            if (transform != null) 
                resultSolid = SolidUtils.CreateTransformed(resultSolid, transform);

            return resultSolid;
        }
    }
}
