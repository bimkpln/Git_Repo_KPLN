using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Looker.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Controls;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Looker.Services
{
    /// <summary>
    /// Сервис по генерации списка пересекаемых элементов
    /// </summary>
    internal static class IntersectedElemFinderService
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
        /// Локальный сервис работы с БД
        /// </summary>
        public static DBWorkerService IntersecedServiceDBWorkerService { get; set; }

        /// <summary>
        /// Метка необходимости анализа проекта на коллизии
        /// </summary>
        public static bool IsDocumentAnalyzing { get; private set; } = false;

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
        /// Список IntersectElemLinkEntity для элементов внутри модели ИОС, которые обрабатываются
        /// </summary>
        public static List<IntersectElemDocEntity> IntersectElemDocEntities { get; private set; }

        /// <summary>
        /// Список IntersectElemLinkEntity для линков ИОС, которые обрабатываются
        /// </summary>
        public static List<IntersectElemDocEntity> IntersectElemLinkEntities { get; private set; }

        public static Document CurrentRevitDoc { get; private set; }

        public static int CurrentDocDBSubDepartmentId { get; private set; }

        public static bool CheckIf_OVLoad { get; set; } = false;

        public static bool CheckIf_VKLoad { get; set; } = false;

        public static bool CheckIf_EOMLoad { get; set; } = false;

        public static bool CheckIf_SSLoad { get; set; } = false;

        /// <summary>
        /// Триггерный метод на обновление данных по текущему проекту. Привязывать к ключевому событию, чтобы данные не протухали. Оптимально - к ViewActivated
        /// </summary>
        public static void UpdateDataByCurrentDocument(Document doc)
        {
#if Revit2020 || Revit2023
            if (CurrentRevitDoc == null || CurrentRevitDoc.Title != doc.Title)
#endif
            if (true)
            {
                CurrentRevitDoc = doc;
                CurrentDocDBSubDepartmentId = IntersecedServiceDBWorkerService.Get_DBDocumentSubDepartmentId(CurrentRevitDoc);

                // ВОТ ЭТО ВСЁ ЗАМЕНИТЬ НА ЧТЕНИЕ ИЗ БД, ПОКА ДЕЛАЮ ЗАГЛУШКУ ЧЕРЕЗ ТХТ ДЛЯ ОПЕРАТИВНОГО ОТКЛЮЧЕНИЯ СЛЕЖЕНИЯ
                FileInfo fi = new FileInfo(@"X:\BIM\5_Scripts\Git_Repo_KPLN\KPLN_Looker\KPLN_Looker\listOfListenenProjects.txt");
                if (fi.Exists)
                {
                    using (StreamReader sr = fi.OpenText())
                    {
                        string content = sr.ReadToEnd();
                        if(CurrentRevitDoc.Title.Split('_').Any(spl => content.Split('~').Contains(spl)))
                            IsDocumentAnalyzing = true;
                        else
                            IsDocumentAnalyzing = false;
                    }
                }
            }
        }


        /// <summary>
        /// Обновить коллекцию сущностей (элементы) для анализа по ТЕКУЩЕМУ проекту по активному виду
        /// </summary>
        public static void UpdateIntersectElemDocEntities(Document doc)
        {
            UIDocument uidoc = new UIDocument(doc);
            View activeView = uidoc.ActiveView;
            if (activeView == null)
            {
                TaskDialog.Show("Ошибка", "Отправь разработчику - не удалось определить класс View");
                return;
            }

            IntersectElemDocEntities = new List<IntersectElemDocEntity>() { GeIntersectElemEntityForDoc(doc, activeView, null) };
        }

        /// <summary>
        /// Обновить коллекцию сущностей (элементы) для анализа по ЛИНКАМ проекта по активному виду
        /// </summary>
        public static void UpdateIntersectElemLinkEntitiesByCurrentView(Document doc)
        {
            // !!!!ТУТ ОШИБКА - НУЖНО ИСПРАВИТЬ Т.К. АНАЛИЗ НАКИНУТ НА ОБНОВЛЕНИЕ ПО ВИДУ, ПО ФАКТУ НУЖНО ОБНОВЛЯТЬ ТОЛЬКО ЭЛЕМЕНЕТЫ ЛИНКОВ В СПЕЦ КОЛЛЕКЦИЮ!!!!
            UIDocument uidoc = new UIDocument(doc);
            View activeView = uidoc.ActiveView;
            if (activeView == null)
            {
                TaskDialog.Show("Ошибка", "Отправь разработчику - не удалось определить класс View");
                return;
            }

            Transform addedDocTrans = doc.ActiveProjectLocation.GetTotalTransform();
            IntersectElemLinkEntities = new List<IntersectElemDocEntity>();

            DocumentSet docSet = doc.Application.Documents;
            foreach (Document openDoc in docSet)
            {
                if (!openDoc.IsLinked || openDoc.Title == doc.Title)
                    continue;
                try
                {
                    // Анализирую модели ИОС разделов на пересечение с создаваемыми
                    int openDocPrjDBSubDepartmentId = IntersecedServiceDBWorkerService.Get_DBDocumentSubDepartmentId(openDoc);
                    switch (openDocPrjDBSubDepartmentId)
                    {
                        case 4:
                            CheckIf_OVLoad = true;
                            goto case 7;
                        case 5:
                            CheckIf_VKLoad = true;
                            goto case 7;
                        case 6:
                            CheckIf_EOMLoad = true;
                            goto case 7;
                        case 7:
                            CheckIf_SSLoad = true;
                            Transform checkTrans = openDoc.ActiveProjectLocation.GetTotalTransform();
                            // Ищу результирующий Transform. Inverse родной, чтобы выяснить разницу
                            Transform resultTransform = addedDocTrans.Inverse * checkTrans;
                            IntersectElemLinkEntities.Add(GeIntersectElemEntityForDoc(openDoc, activeView, resultTransform));
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Print($"Ошибка: {ex.Message}", MessageType.Error);
                }
            }
        }

        /// <summary>
        /// Подготовка коллекции точек пересечений с добавленными эл-тами
        /// </summary>
        public static List<XYZ> GetIntersectPoints(List<Element> addedLinearElems)
        {
            List<XYZ> intersectedPoints = new List<XYZ>();

            foreach (Element addedElem in addedLinearElems)
            {
                intersectedPoints.AddRange(GetIntersectPointsFromCurrentEtities(addedElem, IntersectElemDocEntities));
                intersectedPoints.AddRange(GetIntersectPointsFromCurrentEtities(addedElem, IntersectElemLinkEntities));
            }

            return intersectedPoints;
        }

        /// <summary>
        /// Анализ коллекций спец. классов
        /// </summary>
        /// <returns></returns>
        private static List<XYZ> GetIntersectPointsFromCurrentEtities(Element addedElem, List<IntersectElemDocEntity> checkColl)
        {
            List<XYZ> intersectedPoints = new List<XYZ>();

            Transform addedDocTrans = addedElem.Document.ActiveProjectLocation.GetTotalTransform();

            BoundingBoxXYZ addedElemBB = addedElem.get_BoundingBox(null);
            if (addedElemBB == null)
            {
                HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением BoundingBoxXYZ. " +
                    $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                return intersectedPoints;
            }

            Solid addedElemSolid = GetElemSolid(addedElem);
            if (addedElemSolid == null)
            {
                HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                    $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);
                return intersectedPoints;             }

            foreach (IntersectElemDocEntity docEnt in checkColl)
            {
                Document checkDoc = docEnt.CurrentDoc;
                Transform checkTrans = checkDoc.ActiveProjectLocation.GetTotalTransform();

                // Ищу результирующий Transform. Inverse родной, чтобы выяснить разницу
                Transform resultTransform = addedDocTrans.Inverse * checkTrans;
                Element[] potentialIntersectElems = GetPotentioalIntersectedElems(docEnt, resultTransform, addedElemBB);
                foreach (Element elemToCheck in potentialIntersectElems)
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
        private static Element[] GetPotentioalIntersectedElems(IntersectElemDocEntity docEnt, Transform checkDocTransform, BoundingBoxXYZ addedElemBB)
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
                    addedElemBBMinPnt.X < 0 ? addedElemBBMinPnt.X + expandValue : addedElemBBMinPnt.X - expandValue,
                    addedElemBBMinPnt.Y < 0 ? addedElemBBMinPnt.Y + expandValue : addedElemBBMinPnt.Y - expandValue,
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

            BoundingBoxIntersectsFilter bboxIntersectFilter = new BoundingBoxIntersectsFilter(filterOutline, 0.1);
            BoundingBoxIsInsideFilter bboxInsideFilter = new BoundingBoxIsInsideFilter(filterOutline, 0.1);

            // Подготовка коллекции эл-в пересекаемых и внутри расширенного BoundingBoxXYZ
            potentialIntersectedElems.AddRange(docEnt.CurrenDocPotentialIntersectElemColl
                .Where(e => bboxIntersectFilter.PassesFilter(docEnt.CurrentDoc, e.Id)));

            potentialIntersectedElems.AddRange(docEnt.CurrenDocPotentialIntersectElemColl
                .Where(e => bboxInsideFilter.PassesFilter(docEnt.CurrentDoc, e.Id)));

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

        /// <summary>
        /// Сгенерировать IntersectElemDocEntity для документа по указнному фильтру
        /// </summary>
        private static IntersectElemDocEntity GeIntersectElemEntityForDoc(Document checkDoc, View activeView, Transform resultTrans)
        {
            HashSet<Element> elPotIntColl = new HashSet<Element>(new ElementComparerById());

            Outline filterOutline = CreateViewOutline(resultTrans, activeView);
            foreach (BuiltInCategory category in _builtInCategories)
            {
                if (filterOutline == null)
                {
                    elPotIntColl.UnionWith(new FilteredElementCollector(checkDoc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToElements());
                }
                else
                {
                    elPotIntColl.UnionWith(new FilteredElementCollector(checkDoc)
                        .OfCategory(category)
                        .WherePasses(new BoundingBoxIsInsideFilter(filterOutline, 0.1))
                        .ToElements());

                    elPotIntColl.UnionWith(new FilteredElementCollector(checkDoc)
                        .OfCategory(category)
                        .WherePasses(new BoundingBoxIntersectsFilter(filterOutline, 0.1))
                        .ToElements());
                }
            }

            return new IntersectElemDocEntity(checkDoc, elPotIntColl);
        }

        private static Outline CreateViewOutline(Transform resultTrans, View activeView)
        {
            double expandValue = 10;
            
            BoundingBoxXYZ viewCropBox= activeView.CropBox;
            
            XYZ viewCropMin = viewCropBox.Min;
            XYZ viewCropMax = viewCropBox.Max;
            
            double minZ = 0;
            double maxZ = 0;
            
            Level genLvl = activeView.GenLevel;
            if (genLvl == null)
            {
                minZ = -1000;
                maxZ = +1000;
            }
            else
            {
                minZ = genLvl.Elevation - expandValue;
                maxZ = genLvl.Elevation + expandValue;
            }

            if (!viewCropMin.IsAlmostEqualTo(new XYZ(-100, -100, -1000), 0.1))
            {
                // Подготовка расширенного BoundingBoxXYZ, чтобы не упустить эл-ты
                BoundingBoxXYZ expandedCropBB = new BoundingBoxXYZ()
                {
                    Max = new XYZ(
                    viewCropMax.X < 0 ? viewCropMax.X - expandValue : viewCropMax.X + expandValue,
                    viewCropMax.Y < 0 ? viewCropMax.Y - expandValue : viewCropMax.Y + expandValue,
                    maxZ),
                    Min = new XYZ(
                    viewCropMin.X < 0 ? viewCropMin.X + expandValue : viewCropMin.X - expandValue,
                    viewCropMin.Y < 0 ? viewCropMin.Y + expandValue : viewCropMin.Y - expandValue,
                    minZ),
                };

                Outline filterOutline = null;
                if (resultTrans == null)
                    filterOutline = new Outline(expandedCropBB.Min, expandedCropBB.Max);
                else
                {
                    filterOutline = new Outline(
                        resultTrans.OfPoint(expandedCropBB.Min),
                        resultTrans.OfPoint(expandedCropBB.Max));
                }
                
                if (filterOutline.IsEmpty)
                {
                    XYZ transExpandedElemBBMin = filterOutline.MinimumPoint;
                    XYZ transExpandedElemBBMax = filterOutline.MaximumPoint;

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

                return filterOutline;
            }

            return null;
        }
    }
}
