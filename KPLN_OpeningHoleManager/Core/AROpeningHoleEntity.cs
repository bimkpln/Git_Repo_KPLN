using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.Services;
using KPLN_ModelChecker_Lib.Services.GripGeom.Core;
using KPLN_OpeningHoleManager.Core.MainEntity;
using KPLN_OpeningHoleManager.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    /// <summary>
    /// Кэширование данных по переркрытиям для анализа на расширение границ
    /// </summary>
    internal struct FloorBinding
    {
        internal Document Doc;
        internal Transform Trans;
        internal List<Element> FloorColl;
    }
    
    /// <summary>
    /// Сущность отверстия АР в модели
    /// </summary>
    internal sealed class AROpeningHoleEntity : OpeningHoleEntity
    {
        /// <summary>
        /// Имя параметр для исключения эл-нтов из анализа
        /// </summary>
        internal static string AR_OHE_ParamNameCancelOverwrite = "МО_Перезаписать границы видимости";
        private bool _ar_OHE_IsHostElementKR = false;

        internal AROpeningHoleEntity(Element elem) : base(elem)
        {
        }

        internal AROpeningHoleEntity(Element elem, OpenigHoleShape shape, string iosSubDepCode, Element hostElem, XYZ point) : this(elem)
        {
            OHE_Shape = shape;
            AR_OHE_IOSDubDepCode = iosSubDepCode;
            AR_OHE_HostElement = hostElem;
            OHE_Point = point;
        }

        /// <summary>
        /// Элемент-основа для отверстия
        /// </summary>
        internal Element AR_OHE_HostElement { get; private set; }

        /// <summary>
        /// Элемент-основа для отверстия является элементом КР
        /// </summary>
        internal bool AR_OHE_IsHostElementKR
        {
            get
            {
                if (AR_OHE_HostElement != null)
                {
                    if (AR_OHE_HostElement is Wall wall)
                    {
                        if (wall.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("00_"))
                            _ar_OHE_IsHostElementKR = true;
                    }
                }

                return _ar_OHE_IsHostElementKR;
            }
        }

        /// <summary>
        /// Код раздела, от которого падает задание
        /// </summary>
        internal string AR_OHE_IOSDubDepCode { get; private set; }


        #region Блок пар-в для расширения границ
        /// <summary>
        /// Имя параметра расширения ручек отверстия ВВЕРХ
        /// </summary>
        internal string AR_OHE_VisibilityHandles_ParamNameUpExpander { get; private set; }

        /// <summary>
        /// Привязка к перекрытию выше
        /// </summary>
        public Element UpFloorBinding { get; set; }

        /// <summary>
        /// Расстояние до перекрытия выше
        /// </summary>
        public double UpFloorDistance { get; set; } = double.MaxValue;

        /// <summary>
        /// Имя параметра расширения ручек отверстия ВНИЗ
        /// </summary>
        internal string AR_OHE_VisibilityHandles_ParamNameDownExpander { get; private set; }

        // <summary>
        /// Привязка к перекрытию ниже
        /// </summary>
        public Element DownFloorBinding { get; set; }

        /// <summary>
        /// Расстояние до перекрытия ниже
        /// </summary>
        public double DownFloorDistance { get; set; } = double.MaxValue;
        #endregion

        /// <summary>
        /// Создать объединённое отверстие по выбранной коллекции
        /// </summary>
        internal static AROpeningHoleEntity[] CreateUnionOpeningHole(Document doc, AROpeningHoleEntity[] arOHEColl)
        {
            List<AROpeningHoleEntity> resultColl = new List<AROpeningHoleEntity>();

            // Анализирую на наличие нескольких основ у выборки отверстий
            IEnumerable<int> arOHEHostId = arOHEColl.Select(ohe => ohe.AR_OHE_HostElement.Id.IntegerValue).Distinct();
            foreach (int hostIntId in arOHEHostId)
            {
                // Получаю основу
                Element hostElem = doc.GetElement(new ElementId(hostIntId));

                // Фильтрую коллекцию по основе (объединяю вокруг основы)
                AROpeningHoleEntity[] arOHECollByHost = arOHEColl.Where(ohe => ohe.AR_OHE_HostElement.Id.IntegerValue == hostIntId).ToArray();

                // Подбираю результирующий тип для семейства
                string resultSubDep = string.Empty;
                IEnumerable<string> arOHESubDeps = arOHECollByHost.Select(ohe => ohe.AR_OHE_IOSDubDepCode);
                if (arOHESubDeps.Distinct().Count() > 1 || arOHESubDeps.All(subDep => subDep.Equals("Несколько категорий")))
                    resultSubDep = "Несколько категорий";
                else
                    resultSubDep = arOHESubDeps.FirstOrDefault();


                // Анализирую сущности и нахожу результирующий размер
                Solid unionSolid = null;
                foreach (AROpeningHoleEntity arOHE in arOHECollByHost)
                {
                    try
                    {
                        if (unionSolid == null)
                            unionSolid = arOHE.IGDSolid;
                        else
                        {
                            Solid tempUnionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(unionSolid, arOHE.IGDSolid, BooleanOperationsType.Union);
                            if (tempUnionSolid != null && tempUnionSolid.Volume > 0)
                                unionSolid = tempUnionSolid;

                        }
                    }
                    // Могут быть проблемы с тем, что нельзя выполнить операцию.
                    // Игнорим (возможно стоит добавить отправку пользователю инфы, что такую то стену нужно проверить вручную)
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException) { continue; }
                    catch (Exception ex) { throw ex; }
                }

                // Анализирую вектор основы
                XYZ hostDir = GeometryCurrentWorker.GetHostDirection(hostElem);

                // Получаю ширину и высоту
                double[] widthAndHeight = GeometryCurrentWorker.GetSolidWidhtAndHeight_ByDirection(unionSolid, hostDir);
                double resultWidth = widthAndHeight[0];
                double resultHeight = widthAndHeight[1];


                // Создаю точку вставки
                XYZ unionSolidCentroid = unionSolid.ComputeCentroid();
                BoundingBoxXYZ unionSolidBBox = unionSolid.GetBoundingBox();
                Transform bboxTrans = unionSolidBBox.Transform;

                double locPointX = (bboxTrans.OfPoint(unionSolidBBox.Min).X + bboxTrans.OfPoint(unionSolidBBox.Max).X) / 2;
                double locPointY = (bboxTrans.OfPoint(unionSolidBBox.Min).Y + bboxTrans.OfPoint(unionSolidBBox.Max).Y) / 2;
                double locPointZ = bboxTrans.OfPoint(unionSolidBBox.Min).Z;
                XYZ locPoint = new XYZ(locPointX, locPointY, locPointZ);


                // Создаю сущность для заполнения
                AROpeningHoleEntity resultByHost = new AROpeningHoleEntity(
                    null,
                    OpenigHoleShape.Rectangular,
                    resultSubDep,
                    hostElem,
                    locPoint);

                resultByHost.SetFamilyPathAndName(doc);
                resultByHost.SetGeomParams();
                resultByHost.SetGeomParamsRoundData(resultHeight, resultWidth, 0);

                resultColl.Add(resultByHost);
            }

            return resultColl.ToArray();
        }

        /// <summary>
        /// Очистить коллекцию от ложных инстансов по пересечению
        /// </summary>
        internal static AROpeningHoleEntity[] GetEntitesToDel_ByIntescect(Document doc, AROpeningHoleEntity[] arOHEColl)
        {
            List<AROpeningHoleEntity> elemToClearColl = new List<AROpeningHoleEntity>();

            foreach (AROpeningHoleEntity unionOHEEnt in arOHEColl)
            {
                AROpeningHoleEntity[] AROHEEntColl = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(el => el is FamilyInstance fi
                        && (fi.Symbol.FamilyName.Contains(unionOHEEnt.OHE_FamilyName_Rectangle) || fi.Symbol.FamilyName.Contains(unionOHEEnt.OHE_FamilyName_Circle))
                        && fi.Host.Id.IntegerValue == unionOHEEnt.AR_OHE_HostElement.Id.IntegerValue)
                    .Select(el => new AROpeningHoleEntity(el))
                    .ToArray();

                foreach (AROpeningHoleEntity AROHEEnt in AROHEEntColl)
                {
                    Solid hostOHESolid = GeometryWorker.GetRevitElemUniontSolid(AROHEEnt.IEDElem);
                    if (hostOHESolid == null)
                        continue;


                    if ((unionOHEEnt.IEDElem != null && AROHEEnt.IEDElem != null)
                        && (unionOHEEnt.IEDElem.Id.IntegerValue == AROHEEnt.IEDElem.Id.IntegerValue))
                        continue;

                    //var centr1 = unionOHEEnt.OHE_Solid.ComputeCentroid();
                    //var bbox1 = unionOHEEnt.OHE_Solid.GetBoundingBox();
                    //var trans1 = bbox1.Transform;
                    //var newbbox1 = new BoundingBoxXYZ() { Min = trans1.OfPoint(bbox1.Min), Max = trans1.OfPoint(bbox1.Max)};

                    //var idd2 = checkHostOHE.OHE_Element.Id;
                    //var centr2 = checkHostOHE.OHE_Solid.ComputeCentroid();
                    //var bbox2 = checkHostOHE.OHE_Solid.GetBoundingBox();
                    //var trans2 = bbox2.Transform;
                    //var newbbox2 = new BoundingBoxXYZ() { Min = trans2.OfPoint(bbox2.Min), Max = trans2.OfPoint(bbox2.Max) };

                    Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(hostOHESolid, unionOHEEnt.IGDSolid, BooleanOperationsType.Intersect);
                    if (intersectSolid != null
                        && intersectSolid.Volume > 0
                        && !elemToClearColl.Any(ohe => ohe.IEDElem.Id.IntegerValue == AROHEEnt.IEDElem.Id.IntegerValue))
                        elemToClearColl.Add(AROHEEnt);
                }
            }

            return elemToClearColl.ToArray();
        }

        /// <summary>
        /// Очистка передаваемой коллекции от отверстий, в объединённых стенах (их не нужно создавать, отверстие распространиться на обе стены)
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="arEntities"></param>
        /// <returns></returns>
        internal static AROpeningHoleEntity[] ClearCollectionByJoinedHosts(Document doc, AROpeningHoleEntity[] arEntities)
        {
            List<AROpeningHoleEntity> clearedResult = new List<AROpeningHoleEntity>();

            foreach (AROpeningHoleEntity arEntity in arEntities)
            {
                int[] joinedHostElemIds = JoinGeometryUtils.GetJoinedElements(doc, arEntity.AR_OHE_HostElement).Select(elId => elId.IntegerValue).ToArray();
                if (joinedHostElemIds.Any())
                {
                    AROpeningHoleEntity[] arEntitiesWithJoinedHost = arEntities
                        .Where(ent => joinedHostElemIds.Contains(ent.AR_OHE_HostElement.Id.IntegerValue))
                        .ToArray();

                    if (arEntitiesWithJoinedHost.Any())
                    {
                        // Нахожу ближайшую по отметке Z сущность
                        double tempDistance = 0.1;
                        double arEntityZElev = Math.Round(arEntity.OHE_Point.Z, 3);
                        List<AROpeningHoleEntity> almostEqualZElevColl = new List<AROpeningHoleEntity>();
                        foreach (AROpeningHoleEntity arEntityFRomJoined in arEntitiesWithJoinedHost)
                        {
                            double checkEntZElev = Math.Round(arEntityFRomJoined.OHE_Point.Z, 3);
                            double distance = Math.Abs(arEntityZElev - checkEntZElev);
                            if (distance < tempDistance)
                                almostEqualZElevColl.Add(arEntityFRomJoined);
                        }

                        AROpeningHoleEntity joinedEqualAREnt = null;
                        foreach (AROpeningHoleEntity almostEqualZElev in almostEqualZElevColl)
                        {
                            AROpeningHoleEntity checkedAREnt = joinedEqualAREnt ?? arEntity;

                            // Готовлю индексы имён (гарантированы, т.к. забираю основания только из списка ARKRElemsWorker.ARKRNames_StartWith)
                            string arEntHostIndexName = checkedAREnt.AR_OHE_HostElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Split('_').FirstOrDefault();
                            string almostEqualZElevHostIndexName = almostEqualZElev.AR_OHE_HostElement.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Split('_').FirstOrDefault();

                            // Выбираю сущность по приоритету объединения (ТОЛЬКО ДЛЯ СТЕН)
                            if (checkedAREnt.AR_OHE_HostElement is Wall
                                && almostEqualZElev.AR_OHE_HostElement is Wall
                                && !arEntHostIndexName.Equals(almostEqualZElevHostIndexName))
                            {
                                if (int.TryParse(arEntHostIndexName, out int arEntHostIndex)
                                    && int.TryParse(almostEqualZElevHostIndexName, out int almostEqualZElevHostIndex))
                                    joinedEqualAREnt = arEntHostIndex < almostEqualZElevHostIndex ? checkedAREnt : almostEqualZElev;
                            }
                            // Выбираю сущность по более ТОЛСТОЙ стене
                            else
                            {
                                joinedEqualAREnt = checkedAREnt.AR_OHE_HostElement.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble()
                                    > almostEqualZElev.AR_OHE_HostElement.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED).AsDouble()
                                    ? checkedAREnt
                                    : almostEqualZElev;
                            }
                        }
                        // Убираю дубликаты по координатам (на текущий момент подходят лучше всего)
                        if (joinedEqualAREnt != null && !clearedResult.Any(clEnt => clEnt.OHE_Point.IsAlmostEqualTo(joinedEqualAREnt.OHE_Point, 0.01)))
                            clearedResult.Add(joinedEqualAREnt);
                    }
                    else
                        clearedResult.Add(arEntity);
                }
                else
                    clearedResult.Add(arEntity);
            }

            return clearedResult.ToArray();
        }

        /// <summary>
        /// Обновление файла, чтобы присвоить солид
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="arEntitiesToSet"></param>
        internal static void RegenerateDocAndSetSolids(Document doc, IEnumerable<AROpeningHoleEntity> arEntitiesToSet)
        {
            doc.Regenerate();

            foreach (AROpeningHoleEntity ent in arEntitiesToSet)
            {
                ent.IGDSolid = GeometryWorker.GetRevitElemUniontSolid(ent.IEDElem);
            }
        }

        /// <summary>
        /// Установить путь к Revit семействам
        /// </summary>
        internal AROpeningHoleEntity SetFamilyPathAndName(Document doc)
        {
            if (doc.Title.Contains("СЕТ_1"))
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие прямоугольное.rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\8_Библиотека семейств Самолета\01_АР\Отверстия, шахты, решетки\ASML_АР_Отверстие круглое.rfa";
            }
            else
            {
                OHE_FamilyPath_Rectangle = @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие прямоугольное_(Об_Стена).rfa";
                OHE_FamilyPath_Circle = @"X:\BIM\3_Семейства\1_АР\199_Отверстия, ниши и шахты\199_Отверстие круглое_(Об_Стена).rfa";
            }

            OHE_FamilyName_Rectangle = OHE_FamilyPath_Rectangle
                .Split('\\')
                .FirstOrDefault(part => part.Contains(".rfa"))
                .TrimEnd(".rfa".ToCharArray());

            OHE_FamilyName_Circle = OHE_FamilyPath_Circle
                .Split('\\')
                .FirstOrDefault(part => part.Contains(".rfa"))
                .TrimEnd(".rfa".ToCharArray());

            return this;
        }

        /// <summary>
        /// Установить основные геометрические параметры (ширина, высота, диамтер)
        /// </summary>
        internal AROpeningHoleEntity SetGeomParams()
        {
            if (OHE_Shape == OpenigHoleShape.Rectangular)
            {
                if (OHE_FamilyName_Rectangle.Contains("199_Отверстие прямоугольное"))
                {
                    OHE_ParamNameHeight = "Высота";
                    OHE_ParamNameWidth = "Ширина";
                    OHE_ParamNameExpander = "Расширение границ";
                }
                else if (OHE_FamilyName_Rectangle.Contains("ASML_АР_Отверстие прямоугольное"))
                {
                    OHE_ParamNameHeight = "АР_Высота Проема";
                    OHE_ParamNameWidth = "АР_Ширна Проема";
                }
            }
            else
            {
                if (OHE_FamilyName_Circle.Contains("199_Отверстие круглое"))
                {
                    OHE_ParamNameRadius = "КП_Р_Высота";
                    OHE_ParamNameExpander = "Расширение границ";
                }
                else if (OHE_FamilyName_Circle.Contains("ASML_АР_Отверстие круглое"))
                {
                    OHE_ParamNameRadius = "АР_Высота Проема";
                }
            }

            return this;
        }

        /// <summary>
        /// Установить основные геометрические параметры (ширина, высота, диамтер) И ОКРГУЛИТЬ с шагом 50 мм
        /// </summary>
        internal AROpeningHoleEntity SetGeomParamsRoundData(double height, double width, double radius, double expandValue = 0)
        {
#if Debug2020 || Revit2020
            double roundHeight = RoundGeomParam(height) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
            double roundWidh = RoundGeomParam(width) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
            double roundRadius = RoundGeomParam(radius) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, DisplayUnitType.DUT_MILLIMETERS));
#else
            double roundHeight = RoundGeomParam(height) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
            double roundWidh = RoundGeomParam(width) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
            double roundRadius = RoundGeomParam(radius) + RoundGeomParam(UnitUtils.ConvertToInternalUnits(expandValue, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1")));
#endif

            if (OHE_Shape == OpenigHoleShape.Rectangular)
            {
                if (OHE_FamilyName_Rectangle.Contains("199_Отверстие прямоугольное"))
                {
                    OHE_Height = roundHeight;
                    OHE_Width = roundWidh;
                }
                else if (OHE_FamilyName_Rectangle.Contains("ASML_АР_Отверстие прямоугольное"))
                {
                    OHE_Height = roundHeight;
                    OHE_Width = roundWidh;
                }
            }
            else
            {
                if (OHE_FamilyName_Circle.Contains("199_Отверстие круглое"))
                    OHE_Radius = roundRadius;
                else if (OHE_FamilyName_Circle.Contains("ASML_АР_Отверстие круглое"))
                    OHE_Radius = roundRadius;
            }

            return this;
        }

        /// <summary>
        /// Установить основные геометрические параметры ручек расширения (вверх и вниз)
        /// </summary>
        internal AROpeningHoleEntity SetGeomVisibilityHandlesExpandParams()
        {
            if (OHE_FamilyName_Rectangle.Contains("199_Отверстие прямоугольное"))
            {
                AR_OHE_VisibilityHandles_ParamNameUpExpander = "SYS_OFFSET_UP";
                AR_OHE_VisibilityHandles_ParamNameDownExpander = "SYS_OFFSET_DOWN";
            }
            else if (OHE_FamilyName_Rectangle.Contains("ASML_АР_Отверстие прямоугольное"))
            {
                AR_OHE_VisibilityHandles_ParamNameUpExpander = "АР_Проем видимость вверх";
                AR_OHE_VisibilityHandles_ParamNameDownExpander = "АР_Проем видимость вниз";
            }

            return this;
        }

        /// <summary>
        /// Установить привязки для указания данных в ручки расширения
        /// </summary>
        internal AROpeningHoleEntity SetFloorBindings(Document doc)
        {
            List<FloorBinding> floorBindings = new List<FloorBinding>
            {
                new FloorBinding()
                {
                    Doc = doc,
                    Trans = null,
                    FloorColl = GetFloorsFromDocument(doc),
            }
            };

            // Проход по эл. линков
            var linkInstances = GetARLinkInstances(doc);
            foreach (var rli in linkInstances)
            {
                Document lDoc = rli.GetLinkDocument();
                Transform lTrans = rli.GetTotalTransform();
                List<Element> linkFloors = GetFloorsFromDocument(lDoc, lTrans);
                if (!linkFloors.Any())
                    continue;

                floorBindings.Add(new FloorBinding()
                {
                    Doc = lDoc,
                    Trans = lTrans,
                    FloorColl = linkFloors,
                });

            }

            SetUpFloorBinding(this.IGDGeomCenter, this.IGDBBox.Max, floorBindings);
            SetDownFloorBinding(this.IGDGeomCenter, this.IGDBBox.Min, floorBindings);

            // Если не нашлось - диблированный проход со смещениями в пределах BBox
            if (this.UpFloorBinding == null)
            {
                SetUpFloorBinding(this.IGDBBox.Max, this.IGDBBox.Max, floorBindings);
                SetUpFloorBinding(this.IGDBBox.Min, this.IGDBBox.Max, floorBindings);
            }                
            if (this.DownFloorBinding == null)
            {
                SetDownFloorBinding(this.IGDBBox.Max, this.IGDBBox.Min, floorBindings);
                SetDownFloorBinding(this.IGDBBox.Min, this.IGDBBox.Min, floorBindings);
            }

            return this;
        }

        /// <summary>
        /// Получить перекрытия из текущего документа
        /// </summary>
        private List<Element> GetFloorsFromDocument(Document doc, Transform trans = null)
        {
            // Область поиска
            Outline outline;
            BoundingBoxIntersectsFilter filter;
            if (trans == null)
            {
                outline = GeometryWorker.CreateOutline_ByBBoxANDExpand(this.IGDBBox, new XYZ(0, 0, 30));
                filter = new BoundingBoxIntersectsFilter(outline, 0.1);
            }
            else
            {
                Outline arEntOutline = GeometryWorker.CreateOutline_ByBBoxANDExpand(this.IGDBBox, new XYZ(0, 0, 30));
                outline = GeometryWorker.TransformFilterOutline_ToLink(arEntOutline, trans);
                filter = new BoundingBoxIntersectsFilter(outline, 0.1);
            }

            return new FilteredElementCollector(doc)
                .OfClass(typeof(Floor))
                .WhereElementIsNotElementType()
                .Where(e => ARKRElemsWorker.FloorElemExtraFilterFunc(e) && filter.PassesFilter(e))
                .ToList();
        }

        /// <summary>
        /// Задать привязку к перекрытиям СВЕРХУ
        /// </summary>
        /// <param name="elemCheckPnt">Точка элемента для проверки</param>
        /// <param name="bbMaxPnt">Координаты BBox_Max элемента</param>
        /// <param name="floorColl">Коллекция перекрытий для анализа</param>
        private void SetUpFloorBinding(XYZ elemCheckPnt, XYZ bbMaxPnt, List<FloorBinding> floorBindings)
        {
            foreach (FloorBinding bind in floorBindings)
            {
                if (bind.Trans != null)
                {
                    Transform inv = bind.Trans.Inverse;
                    elemCheckPnt = inv.OfPoint(elemCheckPnt);
                    bbMaxPnt = inv.OfPoint(bbMaxPnt);
                }

                foreach (Element el in bind.FloorColl) 
                { 
                    if (!(el is Floor floor)) continue;

                    XYZ upPrjPoint = floor.GetVerticalProjectionPoint(bbMaxPnt, FloorFace.Bottom);
                    if (upPrjPoint == null) continue;

                    double upDistance = bbMaxPnt.DistanceTo(upPrjPoint);
                    if (Math.Round(bbMaxPnt.Z, 1) <= Math.Round(upPrjPoint.Z, 1) && upDistance <= this.UpFloorDistance)
                    {
                        GeometryElement flGeomElem = floor.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });

                        if (CheckFloorFaces(flGeomElem, new XYZ(elemCheckPnt.X, elemCheckPnt.Y, upPrjPoint.Z)))
                        {
                            this.UpFloorBinding = floor;
                            this.UpFloorDistance = upDistance;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Задать привязку к перекрытиям СВЕРХУ
        /// </summary>
        /// <param name="elemCheckPnt">Точка элемента для проверки</param>
        /// <param name="bbMinPnt">оординаты BBox_Min элемента</param>
        /// <param name="floorColl">Коллекция перекрытий для анализа</param>
        private void SetDownFloorBinding(XYZ elemCheckPnt, XYZ bbMinPnt, List<FloorBinding> floorBindings)
        {
            foreach (FloorBinding bind in floorBindings)
            {
                Transform trans = bind.Trans;
                if (bind.Trans != null)
                {
                    Transform inv = bind.Trans.Inverse;
                    bbMinPnt = inv.OfPoint(bbMinPnt);
                }

                foreach (Element el in bind.FloorColl)
                {
                    if (!(el is Floor floor)) continue;

                    XYZ downPrjPoint = floor.GetVerticalProjectionPoint(bbMinPnt, FloorFace.Top);
                    if (downPrjPoint == null) continue;

                    double downDistance = bbMinPnt.DistanceTo(downPrjPoint);
                    if (Math.Round(bbMinPnt.Z, 1) >= Math.Round(downPrjPoint.Z, 1) && downDistance <= this.DownFloorDistance)
                    {
                        GeometryElement flGeomElem = floor.get_Geometry(new Options() { DetailLevel = ViewDetailLevel.Fine });

                        if (CheckFloorFaces(flGeomElem, new XYZ(elemCheckPnt.X, elemCheckPnt.Y, downPrjPoint.Z)))
                        {
                            this.DownFloorBinding = floor;
                            this.DownFloorDistance = downDistance;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Получить перекрытия из всех линков АР
        /// </summary>
        private static IEnumerable<RevitLinkInstance> GetARLinkInstances(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .Where(l =>
                {
                    string n = l.Name.ToUpper();
                    return n.Contains("_AR_") || n.Contains("_АР_") ||
                           n.Contains("_AR.RVT") || n.Contains("_АР.RVT") ||
                           n.Contains("-AR.RVT") || n.Contains("-АР.RVT") ||
                           n.StartsWith("AR_") || n.StartsWith("АР_");
                });
        }

       /// <summary>
        /// Уточнить значения точки в зависимости от формы
        /// </summary>
        /// <returns></returns>
        internal AROpeningHoleEntity UpdatePointData_ByShape()
        {
            XYZ iosTransPnt = this.OHE_Point;

            if (this.OHE_Shape == OpenigHoleShape.Rectangular)
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - this.OHE_Height / 2);
            else
                OHE_Point = new XYZ(iosTransPnt.X, iosTransPnt.Y, iosTransPnt.Z - this.OHE_Radius / 2);

            return this;
        }

        /// <summary>
        /// Установить солид для СЕТ (там отверстия без тела)
        /// </summary>
        /// <returns></returns>
        internal AROpeningHoleEntity SET_SetSolids()
        {
            this.IGDSolid = GeometryWorker.SET_GetRevitElemUniontSolid(this.IEDElem);

            return this;
        }

        /// <summary>
        /// Разместить экземпляр семейства по указанным координатам и заполнить параметры в модели
        /// </summary>
        internal void CreateIntersectFamInstAndSetRevitParamsData(Document doc)
        {
            // Определяю ключевую часть имени для поиска нужного типа семейства
            string famType = string.Empty;
            if (AR_OHE_IOSDubDepCode.Equals("ОВиК"))
                famType = "ОВ";
            else if (AR_OHE_IOSDubDepCode.Equals("ИТП"))
                famType = "ОВ";
            else if (AR_OHE_IOSDubDepCode.Equals("ВК"))
                famType = "ВК";
            else if (AR_OHE_IOSDubDepCode.Equals("ПТ"))
                famType = "ВК";
            else if (AR_OHE_IOSDubDepCode.Equals("ЭОМ"))
                famType = "ЭОМ";
            else if (AR_OHE_IOSDubDepCode.Equals("СС"))
                famType = "СС";
            else if (AR_OHE_IOSDubDepCode.Equals("АВ"))
                famType = "СС";
            else
                famType = "Несколько категорий";

            FamilySymbol openingFamSymb;
            if (OHE_Shape == OpenigHoleShape.Rectangular)
                openingFamSymb = GetIntersectFamilySymbol(doc, OHE_FamilyPath_Rectangle, OHE_FamilyName_Rectangle, famType);
            else
                openingFamSymb = GetIntersectFamilySymbol(doc, OHE_FamilyPath_Circle, OHE_FamilyName_Circle, famType);

            if (AR_OHE_HostElement.LevelId == null)
                throw new Exception($"У основы с id: {AR_OHE_HostElement.Id} проблемы с привязкой к уровню. Отправь разработчику.");

            Level hostLevel = doc.GetElement(AR_OHE_HostElement.LevelId) as Level;

            // Создание новых экземпляров
            FamilyInstance instance = doc
                .Create
                .NewFamilyInstance(OHE_Point, openingFamSymb, AR_OHE_HostElement, hostLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);


            // Присваиваю параметр эл-та модели инстансу класса (далее используется)
            //doc.Regenerate();
            IEDElem = instance;


            // Указать уровень - для семейств на основе указывать НЕ нужно
            //instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(OHE_Point.Z - hostLevel.Elevation);

            // Заполнить параметры
            if (OHE_Shape == OpenigHoleShape.Rectangular)
            {
                IEDElem.LookupParameter(OHE_ParamNameHeight).Set(OHE_Height);
                IEDElem.LookupParameter(OHE_ParamNameWidth).Set(OHE_Width);

                Parameter expandParam = IEDElem.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null && !expandParam.IsReadOnly)
                    IEDElem.LookupParameter(OHE_ParamNameExpander).Set(0);
            }
            else
            {
                IEDElem.LookupParameter(OHE_ParamNameRadius).Set(OHE_Radius);

                Parameter expandParam = IEDElem.LookupParameter(OHE_ParamNameExpander);
                if (expandParam != null && !expandParam.IsReadOnly)
                    IEDElem.LookupParameter(OHE_ParamNameExpander).Set(0);
            }
        }

        private double RoundGeomParam(double geomParam)
        {
            double round_mm;
#if Debug2020 || Revit2020
            double mm = UnitUtils.ConvertFromInternalUnits(geomParam, DisplayUnitType.DUT_MILLIMETERS);
#else
            double mm = UnitUtils.ConvertFromInternalUnits(geomParam, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
            if (mm % 50 < 0.1)
                round_mm = Math.Round(mm);
            else
                round_mm = Math.Ceiling(mm / 50) * 50;

#if Debug2020 || Revit2020
            return UnitUtils.ConvertToInternalUnits(round_mm, DisplayUnitType.DUT_MILLIMETERS);
#else
            return UnitUtils.ConvertToInternalUnits(round_mm, new ForgeTypeId("autodesk.unit.unit:millimeters-1.0.1"));
#endif
        }

        /// <summary>
        /// Проверка перекрытия на факт пересечения с точкой проецирования
        /// </summary>
        private bool CheckFloorFaces(GeometryElement flGeomElem, XYZ prjPoint)
        {
            foreach (GeometryObject obj in flGeomElem)
            {
                Solid flSolid = obj as Solid;
                if (flSolid != null && flSolid.Volume != 0)
                {
                    FaceArray flFaceArray = flSolid.Faces;
                    foreach (Face flFace in flFaceArray)
                    {
                        IntersectionResult intRes = flFace.Project(prjPoint);
                        if (intRes != null && intRes.Distance < 0.001)
                            return true;
                    }
                }
            }

            return false;
        }
    }
}
