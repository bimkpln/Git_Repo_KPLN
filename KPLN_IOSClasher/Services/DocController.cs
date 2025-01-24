using Autodesk.Revit.DB;
using KPLN_IOSClasher.Core;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_IOSClasher.Services
{
    /// <summary>
    /// Контроллер для подготовки элементов коллизий
    /// </summary>
    internal static class DocController
    {
        /// <summary>
        /// Метка необходимости анализа проекта на коллизии
        /// </summary>
        public static bool IsDocumentAnalyzing { get; private set; }

        /// <summary>
        /// IntersectCheckEntity для элементов внутри модели ИОС, которые обрабатываются
        /// </summary>
        public static IntersectCheckEntity IntersectCheckEntity_Doc { get; private set; }

        /// <summary>
        /// Список IntersectCheckEntity для линков ИОС, которые обрабатываются
        /// </summary>
        public static List<IntersectCheckEntity> IntersectCheckEntity_Link { get; private set; }

        public static Document CheckRevitDoc { get; private set; }

        public static int CheckDocDBSubDepartmentId { get; private set; }

        public static bool CheckIf_OVLoad { get; set; } = false;

        public static bool CheckIf_VKLoad { get; set; } = false;

        public static bool CheckIf_EOMLoad { get; set; } = false;

        public static bool CheckIf_SSLoad { get; set; } = false;

        /// <summary>
        /// Триггерный метод на обновление данных по текущему проекту. Привязывать к ключевому событию, чтобы данные не протухали. Оптимально - к ViewActivated
        /// </summary>
        public static void CurrentDocumentUpdateData(Document doc)
        {
            IsDocumentAnalyzing = true;
            if (doc != null)
            {
                CheckDocDBSubDepartmentId = Module.ModuleDBWorkerService.Get_DBDocumentSubDepartmentId(doc);

                // Глобальный игнор стадии АФК, ПД, АН. Они никогда не проверяются (стадии П+ под вопросом, но чаще всего там только магистрали, пока оставлю так)
                DBProject currentDBPrj = Module.ModuleDBWorkerService.Get_DBProject(doc);
                if (currentDBPrj == null
                    || currentDBPrj.Stage.Equals("АФК") 
                    || currentDBPrj.Stage.Equals("ПД") 
                    || currentDBPrj.Stage.Equals("ПД_Корр") 
                    || currentDBPrj.Stage.Equals("АН"))
                {
                    IsDocumentAnalyzing = false;
                    return;
                }
                
                DBProjectsIOSClashMatrix[] prjMatrix = Module.ModuleDBWorkerService.Get_DBProjectsIOSClashMatrix(currentDBPrj).ToArray();
                // Это - затычка. Когда перейдём на схему контроля по всем проектам, эту часть нужно заменить, т.к. все объекты, кроме
                // тех, на которые есть исключения - игнорируются, а этого быть не должно
                if (prjMatrix.Count() == 0)
                    IsDocumentAnalyzing = false;
                // Игнор по всему отделу
                else if (prjMatrix.Any(prj => prj.ExceptionSubDepartmentId == Module.ModuleDBWorkerService.CurrentDBUser.SubDepartmentId))
                    IsDocumentAnalyzing = false;
                // Игнор по отдельным пользователям
                else
                    IsDocumentAnalyzing = !prjMatrix.Any(prj => prj.ExceptionUserId == Module.ModuleDBWorkerService.CurrentDBUser.Id);
            }

#if Debug2020 || Debug2023
            //IsDocumentAnalyzing = true;
#endif
        }

        /// <summary>
        /// Обновить коллекцию сущностей (элементы) для анализа по ТЕКУЩЕМУ проекту, с учетом добавленных элементов
        /// </summary>
        public static void UpdateIntCheckEntities_Doc(Document doc, IEnumerable<Element> elems)
        {
            if (elems.Any())
            {
                BoundingBoxXYZ filterBBox = CreateElemsBBox(elems);
                Outline filterOutline = CreateFilterOutline(filterBBox, 1);

                IntersectCheckEntity_Doc = new IntersectCheckEntity(doc, filterBBox, filterOutline);
            }
        }

        /// <summary>
        /// Обновить коллекцию сущностей (элементы) для анализа по ЛИНКАМ проекта, с учетом активного окна (если его нет, беру самостоятельно)
        /// </summary>
        public static void UpdateIntCheckEntities_Link(Document doc, View activeView)
        {
            BoundingBoxXYZ filterBBox = CreateViewBBox(activeView);
            if (filterBBox == null)
                return;

            Outline filterOutline = CreateFilterOutline(filterBBox, 1);

            IntersectCheckEntity_Link = new List<IntersectCheckEntity>();

            RevitLinkInstance[] revitLinkInsts = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .WhereElementIsNotElementType()
                .Cast<RevitLinkInstance>()
                .ToArray();

            DocumentSet docSet = doc.Application.Documents;

            // Обнуляю статусы загрузок
            CheckIf_OVLoad = false;
            CheckIf_VKLoad = false;
            CheckIf_EOMLoad = false;
            CheckIf_SSLoad = false;

            foreach (Document openDoc in docSet)
            {
                if (!openDoc.IsLinked || openDoc.Title == doc.Title)
                    continue;
                try
                {
                    // Анализирую модели ИОС разделов на пересечение с создаваемыми
                    int openDocPrjDBSubDepartmentId = Module.ModuleDBWorkerService.Get_DBDocumentSubDepartmentId(openDoc);
                    switch (openDocPrjDBSubDepartmentId)
                    {
                        case 4:
                            CheckIf_OVLoad = true;
                            goto case 99;
                        case 5:
                            CheckIf_VKLoad = true;
                            goto case 99;
                        case 6:
                            CheckIf_EOMLoad = true;
                            goto case 99;
                        case 7:
                            CheckIf_SSLoad = true;
                            goto case 99;
                        case 99:
                            RevitLinkInstance rLink = revitLinkInsts
                                .FirstOrDefault(rl => openDoc.Title.Contains(rl.Name.Split(new string[] { ".rvt" }, StringSplitOptions.None)
                                .FirstOrDefault()));

                            // Если открыто сразу несколько моделей одного проекта, то линки могут прилететь с другого файла. В таком случае - игнор
                            if (rLink != null)
                                IntersectCheckEntity_Link.Add(new IntersectCheckEntity(doc, filterBBox, filterOutline, rLink));

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
        /// Подготовка коллекции классов точек пересечений с добавленными эл-тами
        /// </summary>
        public static HashSet<IntersectPointEntity> GetIntPntEntities(Element[] addedLinearElems)
        {
            HashSet<IntersectPointEntity> intersectedPoints = new HashSet<IntersectPointEntity>();

            foreach (Element addedElem in addedLinearElems)
            {
                BoundingBoxXYZ addedElemBB = addedElem.get_BoundingBox(null);
                if (addedElemBB == null)
                {
                    HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением BoundingBoxXYZ. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);

                    return null;
                }

                Solid addedElemSolid = GetElemSolid(addedElem);
                if (addedElemSolid == null)
                {
                    HtmlOutput.Print($"У элемента с id: {addedElem.Id} из ТВОЕЙ модели проблемы с получением Solid. " +
                        $"Элемент проигнорирован, ошибку стоит отправить разработчику);", MessageType.Error);

                    return null;
                }


                // Анализирую элементы внутри файла
                foreach (Element elemToCheck in IntersectCheckEntity_Doc.CurrentDocElemsToCheck)
                {

                    IntersectPointEntity newEntity = GetPntEntityFromElems(addedElem, addedElemSolid, elemToCheck, IntersectCheckEntity_Doc);
                    if (newEntity == null) continue;

                    intersectedPoints.Add(newEntity);
                }


                // Анализирую элементы в линках
                Outline addedElemOutline = CreateFilterOutline(addedElemBB, 1);

                foreach (IntersectCheckEntity checkEnt in IntersectCheckEntity_Link)
                {
                    Element[] potentialIntersectElems = checkEnt.GetPotentioalIntersectedElems_ForLink(addedElemOutline);

                    foreach (Element elemToCheck in potentialIntersectElems)
                    {
                        IntersectPointEntity newEntity = GetPntEntityFromElems(addedElem, addedElemSolid, elemToCheck, checkEnt);
                        if (newEntity == null) continue;

                        intersectedPoints.Add(newEntity);
                    }
                }
            }

            return intersectedPoints;
        }

        /// <summary>
        /// Создать BoundingBoxXYZ для добавленных элементов
        /// </summary>
        public static BoundingBoxXYZ CreateElemsBBox(IEnumerable<Element> elems)
        {
            BoundingBoxXYZ resultBBox = null;

            foreach (Element element in elems)
            {
                BoundingBoxXYZ elementBox = element.get_BoundingBox(null);
                if (elementBox == null)
                    continue;

                if (resultBBox == null)
                {
                    resultBBox = new BoundingBoxXYZ
                    {
                        Min = elementBox.Min,
                        Max = elementBox.Max
                    };
                }
                else
                {
                    resultBBox.Min = new XYZ(
                        Math.Min(resultBBox.Min.X, elementBox.Min.X),
                        Math.Min(resultBBox.Min.Y, elementBox.Min.Y),
                        Math.Min(resultBBox.Min.Z, elementBox.Min.Z));

                    resultBBox.Max = new XYZ(
                        Math.Max(resultBBox.Max.X, elementBox.Max.X),
                        Math.Max(resultBBox.Max.Y, elementBox.Max.Y),
                        Math.Max(resultBBox.Max.Z, elementBox.Max.Z));
                }
            }

            return resultBBox;
        }

        /// <summary>
        /// Создать Outline по указанному BoundingBoxXYZ и расширению
        /// </summary>
        public static Outline CreateFilterOutline(BoundingBoxXYZ bbox, double expandValue)
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

            Print($"Отправь разработчику - не удалось создать Outline для фильтрации", MessageType.Error);
            return null;
        }

        public static IntersectPointEntity GetPntEntityFromElems(Element addedElem, Solid addedElemSolid, Element elemToCheck, IntersectCheckEntity intEnt)
        {
            if (elemToCheck.Id.IntegerValue == addedElem.Id.IntegerValue)
                return null;

            int linkInstanceId = -1;
            if (intEnt.CheckLinkInst != null)
                linkInstanceId = intEnt.CheckLinkInst.Id.IntegerValue;

            Solid solidToCheck = GetElemSolid(elemToCheck, intEnt.LinkTransfrom);
            Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(addedElemSolid, solidToCheck, BooleanOperationsType.Intersect);
            if (intersectionSolid != null && intersectionSolid.Volume > 0)
                return new IntersectPointEntity(
                    intersectionSolid.ComputeCentroid(),
                    addedElem.Id.IntegerValue,
                    elemToCheck.Id.IntegerValue,
                    linkInstanceId,
                    Module.ModuleDBWorkerService.CurrentDBUser);

            return null;
        }

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
        /// Создать BoundingBoxXYZ вида по ТЕКУЩЕМУ документу
        /// </summary>
        private static BoundingBoxXYZ CreateViewBBox(View activeView)
        {
            BoundingBoxXYZ combinedBoundingBox = null;
            // Если активна подрезка на плане - берем её (учитываем уровень с расширением)
            if (activeView is ViewPlan && activeView.CropBoxActive)
            {
                Level viewLevel = activeView.GenLevel;
                double levelZCoord = viewLevel.Elevation;
                BoundingBoxXYZ vCropBB = activeView.CropBox;
                XYZ vCropBBMin = vCropBB.Min;
                XYZ vCropBBMax = vCropBB.Max;

                return new BoundingBoxXYZ()
                {
                    Min = new XYZ(vCropBBMin.X, vCropBBMin.Y, levelZCoord - 10),
                    Max = new XYZ(vCropBBMax.X, vCropBBMax.Y, levelZCoord + 30),
                };
            }
            // Если активна подрезка на плане - берем её
            else if (activeView is View3D view3D && view3D.IsSectionBoxActive)
            {
                BoundingBoxXYZ sectBox = view3D.GetSectionBox();
                Transform viewTrans = sectBox.Transform;
                return new BoundingBoxXYZ()
                {
                    Min = viewTrans.OfPoint(sectBox.Min),
                    Max = viewTrans.OfPoint(sectBox.Max),
                };
            }
            // Если активна подрезка на разрезе - берем её
            else if (activeView is ViewSection viewSection && activeView.CropBoxActive)
            {
                // В координатах плана
                var cropBox = viewSection.CropBox;
                double sectionDepth = Math.Abs(cropBox.Max.Z - cropBox.Min.Z);

                // В координатах самого окна
                var cropManager = viewSection.GetCropRegionShapeManager();
                IList<CurveLoop> cropShape = cropManager.GetCropShape();
                XYZ viewDirection = viewSection.ViewDirection * sectionDepth;
                cropShape.Add(CurveLoop.CreateViaTransform(cropShape.FirstOrDefault(), Transform.CreateTranslation(viewDirection.Negate())));

                Solid fullCropSolid = GeometryCreationUtilities.CreateLoftGeometry(cropShape, new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId));

                return fullCropSolid.GetBoundingBox();
            }
            else if (activeView is ViewSheet viewSheet)
            {
                ISet<ElementId> vPortIds = viewSheet.GetAllPlacedViews();
                if (!vPortIds.Any()) return null;

                if (!vPortIds
                    .Select(id => activeView.Document.GetElement(id))
                    .Any(el => el is ViewPlan || el is View3D || el is ViewSection)) return null;
            }
            else if (activeView.GenLevel != null)
            {
                Level viewLevel = activeView.GenLevel;
                double levelZCoord = viewLevel.Elevation;
                return new BoundingBoxXYZ()
                {
                    Min = new XYZ(-2000, -2000, levelZCoord - 10),
                    Max = new XYZ(2000, 2000, levelZCoord + 30),
                };
            }

            // Иначе - генерю BoundingBoxXYZ исходя из элементов ТЕКУЩЕГО документа
            Document activeDoc = activeView.Document;
            Element[] elColl = new FilteredElementCollector(activeDoc)
                .WherePasses(IntersectCheckEntity.ElemCatLogicalOrFilter)
                .ToArray();
            foreach (Element element in elColl)
            {
                BoundingBoxXYZ elementBox = element.get_BoundingBox(activeView);
                if (elementBox == null)
                    continue;

                if (combinedBoundingBox == null)
                {
                    combinedBoundingBox = new BoundingBoxXYZ
                    {
                        Min = elementBox.Min,
                        Max = elementBox.Max
                    };
                }
                else
                {
                    combinedBoundingBox.Min = new XYZ(
                        Math.Min(combinedBoundingBox.Min.X, elementBox.Min.X),
                        Math.Min(combinedBoundingBox.Min.Y, elementBox.Min.Y),
                        Math.Min(combinedBoundingBox.Min.Z, elementBox.Min.Z));

                    combinedBoundingBox.Max = new XYZ(
                        Math.Max(combinedBoundingBox.Max.X, elementBox.Max.X),
                        Math.Max(combinedBoundingBox.Max.Y, elementBox.Max.Y),
                        Math.Max(combinedBoundingBox.Max.Z, elementBox.Max.Z));
                }
            }

            return combinedBoundingBox;
        }
    }
}
