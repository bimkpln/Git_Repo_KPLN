using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    /// <summary>
    /// Абстрактный класс строителя 
    /// </summary>
    internal abstract class AbstrGripBuilder
    {
        /// <summary>
        ///  GUID параметра для исключения перезаписи ("ПЗ_Перезаписать")
        /// </summary>
        private readonly Guid _revalueParamGuid = new Guid("38157d2d-f952-41e5-8d05-2c962addfe56");
        private List<InstanceElemData> _elemsOnLevel = new List<InstanceElemData>();
        private List<InstanceElemData> _elemsByHost = new List<InstanceElemData>();
        private List<InstanceElemData> _elemsUnderLevel = new List<InstanceElemData>();
        private List<InstanceElemData> _stairsElems = new List<InstanceElemData>();
        private List<InstanceElemData> _allElements = new List<InstanceElemData>();
        private List<LevelAndSectionSolid> _sectDataSolids = new List<LevelAndSectionSolid>();
        private List<GripParamError> _errorElements = new List<GripParamError>();
        private int _allElementsCount = 0;
        private int _hostElementsCount = 0;

        public AbstrGripBuilder(Document doc, string docMainTitle, string levelParamName, string sectionParamName)
        {
            Doc = doc;
            DocMainTitle = docMainTitle;
            LevelParamName = levelParamName;
            SectionParamName = sectionParamName;
        }

        /// <summary>
        /// Документ Ревит
        /// </summary>
        protected internal Document Doc { get; }

        /// <summary>
        /// Имя документа Ревит
        /// </summary>
        internal string DocMainTitle { get; }

        /// <summary>
        /// Имя параметра, в который осуществляется запись уровня
        /// </summary>
        internal string LevelParamName { get; }

        /// <summary>
        /// Имя параметра, в который осуществляется запись секции
        /// </summary>
        internal string SectionParamName { get; }

        /// <summary>
        /// Счетчик выпроненных операций по записи данных
        /// </summary>
        internal int PbCounter { get; private set; } = 0;

        /// <summary>
        /// Коллекция элементов на уровне
        /// </summary>
        public List<InstanceElemData> ElemsOnLevel
        {
            get => _elemsOnLevel;
            protected set => _elemsOnLevel = value;
        }

        /// <summary>
        /// Коллекция элементов, которые являются общими вложенными
        /// </summary>
        public List<InstanceElemData> ElemsByHost
        {
            get => _elemsByHost;
            protected set => _elemsByHost = value;
        }

        /// <summary>
        /// Коллекция элементов под уровне
        /// </summary>
        public List<InstanceElemData> ElemsUnderLevel
        {
            get => _elemsUnderLevel;
            protected set => _elemsUnderLevel = value;
        }

        /// <summary>
        /// Коллекция лестниц
        /// </summary>
        public List<InstanceElemData> StairsElems
        {
            get => _stairsElems;
            protected set => _stairsElems = value;
        }

        /// <summary>
        /// Коллекция лестниц
        /// </summary>
        public List<InstanceElemData> AllElements
        {
            get
            {
                _allElements = ElemsOnLevel.Concat(ElemsByHost).Concat(ElemsUnderLevel).Concat(StairsElems).ToList();
                return _allElements;
            }
        }

        /// <summary>
        /// Количество всех элементов
        /// </summary>
        public int AllElementsCount
        {
            get
            {
                _allElementsCount = AllElements.Count;
                if (_allElementsCount == 0)
                    throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет, или имя проекта не соответсвует ВЕР!\n");

                return _allElementsCount;
            }
        }

        /// <summary>
        /// Количество элементов с основанием
        /// </summary>
        public int HostElementsCount
        {
            get
            {
                if (_hostElementsCount == 0)
                    _hostElementsCount = ElemsByHost.Count;

                return _hostElementsCount;
            }
        }

        /// <summary>
        /// Коллекция спец. солидов с разделением по уровням
        /// </summary>
        public List<LevelAndSectionSolid> SectDataSolids
        {
            get => _sectDataSolids;
            protected set => _sectDataSolids = value;
        }


        /// <summary>
        /// Коллекция элементов, которые при анализе выдали ошибку
        /// </summary>
        public List<GripParamError> ErrorElements
        {
            get => _errorElements;
            private set => _errorElements = value;
        }

        /// <summary>
        /// Метод подготовки элементов к обработке
        /// </summary>
        public abstract void Prepare();

        /// <summary>
        /// Метод проверки элементов на заполняемость параметров и чистка списков от элементов, не подверженных чистке
        /// </summary>
        public virtual void Check(Document doc)
        {
            Task elemsOnLevelCheckTask = Task.Run(() => CheckElemParams(ElemsOnLevel));
            Task elemsByHostCheckTask = Task.Run(() => CheckElemParams(ElemsByHost));
            Task elemsUnderLevelCheckTask = Task.Run(() => CheckElemParams(ElemsUnderLevel));
            Task elemsStairsElemsCheckTask = Task.Run(() => CheckElemParams(StairsElems));

            ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, AllElements.Select(e => e.IEDElem.Id).ToArray());
            int errorCount = AllElementsCount - availableWSElemsId.Count;
            if (errorCount > 0)
                throw new GripParamExection($"Возможность изменения ограничена для {errorCount} элементов. Попроси коллег ОСВОБОДИТЬ все забранные рабочие наборы и элементы\n");

            Task.WaitAll(new Task[] { elemsOnLevelCheckTask, elemsByHostCheckTask, elemsUnderLevelCheckTask, elemsStairsElemsCheckTask });
        }

        /// <summary>
        /// Метод заполнения парамтеров захваток (секции и уровня) по анализу геометрии
        /// </summary>
        /// <param name="pb">Прогресс-бар для визуализации процесса выполнения</param>
        public void ExecuteGripParams_ByGeom(Progress_Single pb)
        {
            foreach (InstanceElemData instElemData in ElemsOnLevel)
            {
                Parameter instElemDataSectParam = instElemData.IEDElem.LookupParameter(SectionParamName);
                Parameter instElemDataLvlParam = instElemData.IEDElem.LookupParameter(LevelParamName);

                // Если залочен у общего вложенного, то 99%, что это он передаётся из родителя
                if (instElemData.IEDElem is FamilyInstance famInst
                    && famInst.SuperComponent != null
                    && (instElemDataSectParam.IsReadOnly || instElemDataLvlParam.IsReadOnly))
                {
                    ErrorElements.Add(new GripParamError(
                            instElemData.IEDElem,
                            "Блокировка параметра: у общего вложенного семейства параметр для секции или этажа заблокирован. Скорее всего, он передаётся из родителя, но нужно проверить"));
                    continue;
                }

                InstanceGeomData instGeomData = (InstanceGeomData)instElemData
                    ?? throw new GripParamExection(
                        $"Элемент {instElemData.IEDElem.Id} был не правильно назначен (как элемент без геометрии. Обратись к разработчику\n");


                LevelAndSectionSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestMaxIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.IEDElem,
                            "Геометрия: Элементу не удалось присвоить данные по геометрии"));
                        continue;
                    }
                }

                if (instElemDataLvlParam.IsReadOnly || instElemDataSectParam.IsReadOnly)
                    throw new GripParamExection($"У элемента id: {instElemData.IEDElem.Id} заблокирован один из параметров для записи захваток: {LevelParamName}, или {SectionParamName}");

                instElemDataLvlParam.Set(maxIntersectInstance.LSLevelData);
                instElemDataSectParam.Set(maxIntersectInstance.LSSectionData);
                instElemData.IsEmptyData = false;

                pb.Update(++PbCounter, "Поиск по геометрии");
            }

            foreach (InstanceElemData instElemData in ElemsUnderLevel)
            {
                Parameter instElemDataSectParam = instElemData.IEDElem.LookupParameter(SectionParamName);
                Parameter instElemDataLvlParam = instElemData.IEDElem.LookupParameter(LevelParamName);

                // Если залочен у общего вложенного, то 99%, что это он передаётся из родителя
                if (instElemData.IEDElem is FamilyInstance famInst
                    && famInst.SuperComponent != null
                    && (instElemDataSectParam.IsReadOnly || instElemDataLvlParam.IsReadOnly))
                {
                    ErrorElements.Add(new GripParamError(
                            instElemData.IEDElem,
                            "Блокировка параметра: у общего вложенного семейства параметр для секции или этажа заблокирован. Скорее всего, он передаётся из родителя, но нужно проверить"));
                    continue;
                }

                InstanceGeomData instGeomData = (InstanceGeomData)instElemData ??
                    throw new GripParamExection(
                        $"Элемент {instElemData.IEDElem.Id} был не правильно назначен (как элемент без гометриии. Обратись к разработчику");

                LevelAndSectionSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestMaxIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.IEDElem,
                            "Геометрия: Элементу не удалось присвоить данные по геометрии"));
                        continue;
                    }
                }
                LevelAndSectionSolid downLevelAndGridSolid = SectDataSolids
                    .Where(s =>
                        s.LSSectionData.Equals(maxIntersectInstance.LSSectionData)
                        && s.LSLevelData.Equals(maxIntersectInstance.LSLevelData))
                    .FirstOrDefault();

                instElemDataLvlParam.Set(maxIntersectInstance.LSLevelData);
                instElemDataSectParam.Set(maxIntersectInstance.LSSectionData);
                instElemData.IsEmptyData = false;


                pb.Update(++PbCounter, "Поиск по геометрии");
            }
        }

        /// <summary>
        /// Метод заполнения парамтеров захваток (секции и уровня) по анализу оснований. 
        /// ВАЖНО: запускать после анализа геометрии и в другой транзакции
        /// </summary>
        /// <param name="pb">Прогресс-бар для визуализации процесса выполнения</param>
        public void ExecuteGripParams_ByHost(Progress_Single pb)
        {
            if (!ElemsByHost.Any())
            {
                pb.Update(PbCounter, "Анализ элементов на основе");
                return;
            }

            foreach (InstanceElemData instElemData in ElemsByHost)
            {
                Element elem = instElemData.IEDElem;
                Element hostElem = null;
                // Элементы АР - панели
                if (elem is Panel panel)
                    hostElem = (Wall)panel.Host;
                // Элементы АР - витражи
                else if (elem is Mullion mullion)
                    hostElem = (Wall)mullion.Host;
                // Вложенные семейства
                else if (elem is FamilyInstance instance)
                    hostElem = instance.SuperComponent;
                // Изоляция ИОС
                else if (elem is InsulationLiningBase insulationLiningBase)
                    hostElem = Doc.GetElement(insulationLiningBase.HostElementId);

                if (hostElem != null)
                    SetParameter(instElemData, hostElem);
                else
                {
                    ErrorElements.Add(new GripParamError(
                        instElemData.IEDElem,
                        "Вложенность: Не удалось определить основу. Скинь разработчику!"));
                }

                instElemData.IsEmptyData = false;

                pb.Update(++PbCounter, "Анализ элементов на основе");
            }
        }

        /// <summary>
        /// Заполняю данными элементы, котоыре не подверглись обработке
        /// </summary>
        public void CheckNotExecutedElems()
        {
            ErrorElements.AddRange(
                AllElements
                .Where(e => e.IsEmptyData)
                .Select(e => new GripParamError(e.IEDElem, "Элементы не подверглись анализу (это ПОЛНЫЙ список, ниже будут списки с отдельными классификациями)")));
        }

        /// <summary>
        /// Метод проверки и очистки элементов
        /// </summary>
        /// <param name="checkColl">Коллекция для проверки</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private void CheckElemParams(List<InstanceElemData> checkColl)
        {
            checkColl.RemoveAll(instData =>
            {
                Element elem = instData.IEDElem;
                Parameter sectParam = elem.LookupParameter(SectionParamName);
                Parameter levParam = elem.LookupParameter(LevelParamName);
                if (sectParam == null || levParam == null)
                {
                    throw new GripParamExection(
                        $"Прервано по причине отсутствия необходимых параметров захваток (секции или этажа).\n" +
                        $"Пример: Категория: {elem.Category.Name} / id: {elem.Id}");
                }
                else if (elem is FamilyInstance fam && fam.SuperComponent == null && (sectParam.IsReadOnly || levParam.IsReadOnly))
                {
                    throw new GripParamExection(
                        $"Прервано из-за того, что параметр захваток (секции или этажа) заблокирован.\n" +
                        $"Элемен: id: {elem.Id}");
                }

                Parameter canReValueParam = elem.get_Parameter(_revalueParamGuid);

                return canReValueParam != null && canReValueParam.HasValue && canReValueParam.AsInteger() != 1;
            });
        }

        /// <summary>
        /// Определение солида, который пересекается с солидом секции. Солид эл-та Ревит ПРИТЯГИВАЕТСЯ к солиду секции БЕЗ изменения координаты Z
        /// </summary>
        /// <param name="instSolid">Солид эл-та ревит для проверки</param>
        /// <param name="sectData">Солид секции для проверки</param>
        /// <returns></returns>
        private Solid GetNearestHorizontalIntesectedInstSolid(Solid instSolid, LevelAndSectionSolid sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.LSSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(sectTransform.Origin.X, sectTransform.Origin.Y, instTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.LSSolid, BooleanOperationsType.Intersect);
            if (intersectSolid != null && intersectSolid.Volume > 0)
                return intersectSolid;

            return null;
        }

        /// <summary>
        /// Определение солида, который пересекается с солидом секции. Солид эл-та Ревит ПРИТЯГИВАЕТСЯ к солиду секции БЕЗ изменения координат X и Y
        /// </summary>
        /// <param name="instSolid">Солид эл-та ревит для проверки</param>
        /// <param name="sectData">Солид секции для проверки</param>
        /// <returns></returns>
        private Solid GetNearestVerticalIntesectedInstSolid(Solid instSolid, LevelAndSectionSolid sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.LSSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(instTransform.Origin.X, instTransform.Origin.Y, sectTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.LSSolid, BooleanOperationsType.Intersect);
            if (intersectSolid != null && intersectSolid.Volume > 0)
                return intersectSolid;

            return null;
        }

        /// <summary>
        /// Заполнить параметры элемента по данным из его основы
        /// </summary>
        /// <param name="instElemData">Элемент для записи</param>
        /// <param name="hostElem">Хост</param>
        private void SetParameter(InstanceElemData instElemData, Element hostElem)
        {
            string hostElemSectParamValue = hostElem.LookupParameter(SectionParamName).AsString();
            string hostElemLevParamValue = hostElem.LookupParameter(LevelParamName).AsString();
            if (hostElemSectParamValue == null
                || string.IsNullOrEmpty(hostElemSectParamValue)
                || hostElemLevParamValue == null
                || string.IsNullOrEmpty(hostElemLevParamValue))
            {
                ErrorElements.Add(new GripParamError(
                    instElemData.IEDElem,
                    $"Вложенность: У элемента основы (id: {hostElem.Id}) не заполнены данные для передачи"));
            }
            else
            {
                Parameter elemSectParam = instElemData.IEDElem.LookupParameter(SectionParamName);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemSectParam.IsReadOnly)
                    elemSectParam.Set(hostElemSectParamValue);

                Parameter elemLevParam = instElemData.IEDElem.LookupParameter(LevelParamName);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemLevParam.IsReadOnly)
                    elemLevParam.Set(hostElemLevParamValue);
            }
        }

        /// <summary>
        /// Получить LevelAndGridSolid с наибольшим пересечением по солидам
        /// </summary>
        /// <param name="instGeomData">Элемент для анализа</param>
        private LevelAndSectionSolid GetMaxIntersectedLevelAndGridSolid(InstanceGeomData instGeomData)
        {
            LevelAndSectionSolid result = null;
            double maxIntersectValue = 0;

            foreach (LevelAndSectionSolid levelAndGridSolid in SectDataSolids)
            {
                // Фильтры
                BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(levelAndGridSolid.BBoxOutline, 0.1);
                BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(levelAndGridSolid.BBoxOutline, 0.1);


                // Фильтрация по QuickFilter
                if (intersectsFilter.PassesFilter(instGeomData.IEDElem) || insideFilter.PassesFilter(instGeomData.IEDElem))
                {
                    double tempIntersectValue = 0;
                    try
                    {
                        // Проверяю положение в секции
                        Solid checkIntersectSectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            instGeomData.IGDSolid,
                            levelAndGridSolid.LSSolid,
                            BooleanOperationsType.Intersect);

                        if (checkIntersectSectSolid == null || !(checkIntersectSectSolid.Volume > 0))
                            continue;

                        tempIntersectValue += Math.Round(checkIntersectSectSolid.Volume, 10);

                        if (tempIntersectValue > 0 && Math.Round(Math.Abs(tempIntersectValue) - (Math.Abs(maxIntersectValue)), 2) >= 0)
                        {
                            maxIntersectValue = tempIntersectValue;
                            result = levelAndGridSolid;
                        }
                    }
                    // Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        ErrorElements.Add(new GripParamError(
                            instGeomData.IEDElem,
                            "Геометрия: Элемент нужно назначить вручную (геометрию невозможно проанализиовать)"));
                    }
                    catch (Exception ex)
                    {
                        throw new GripParamExection($"Что-то непонятное с элементом с id: {instGeomData.IEDElem.Id}. Отправь разработчику:\n {ex.Message}");
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Получить LevelAndGridSolid, который расположен ближе всего, и который имеет макс. объем по пересечению
        /// </summary>
        /// <param name="instGeomData">Элемент для анализа</param>
        private LevelAndSectionSolid GetNearestMaxIntersectedLevelAndGridSolid(InstanceGeomData instGeomData)
        {
            // Переменные для поиска ближайшей секции
            LevelAndSectionSolid resultLSSolid = null;
            double maxIntersectValue = 0;
            // 4,5 м - условно возможное отклонение элемента от солида
            double minPrjDistanceValue = 15;

            foreach (LevelAndSectionSolid levelAndGridSolid in SectDataSolids)
            {
                // Фильтры
                BoundingBoxIntersectsFilter intersectsFilter = new BoundingBoxIntersectsFilter(levelAndGridSolid.BBoxOutline, minPrjDistanceValue);
                BoundingBoxIsInsideFilter insideFilter = new BoundingBoxIsInsideFilter(levelAndGridSolid.BBoxOutline, minPrjDistanceValue);


                // Фильтрация по QuickFilter
                if (intersectsFilter.PassesFilter(instGeomData.IEDElem) || insideFilter.PassesFilter(instGeomData.IEDElem))
                {
                    double tempIntersectValue = 0;
                    double tempPrjDistanceValue = minPrjDistanceValue * 1.1;
                    try
                    {
                        FaceArray levelAndGridFaceArray = levelAndGridSolid.LSSolid.Faces;

                        #region Првоеряю положение в секции в плоскости XY
                        Solid resSolid = GetNearestHorizontalIntesectedInstSolid(instGeomData.IGDSolid, levelAndGridSolid);
                        if (resSolid != null && resSolid.Volume > 0)
                        {
                            tempIntersectValue += resSolid.Volume;

                            XYZ resSolidCentroid = resSolid.ComputeCentroid();
                            if (resSolidCentroid == null)
                                continue;

                            // Проверяю расстояние до секции с целью выявления ближайшей
                            foreach (Face levelAndGridFace in levelAndGridFaceArray)
                            {
                                IntersectionResult prjPointResult = levelAndGridFace.Project(resSolidCentroid);
                                if (prjPointResult != null && prjPointResult.Distance < tempPrjDistanceValue)
                                    tempPrjDistanceValue = Math.Round(prjPointResult.Distance, 3);
                            }

                            bool checkValue = (tempIntersectValue > 0 && Math.Round((tempIntersectValue - maxIntersectValue), 2) >= 0);
                            bool checkDistanceXY = (Math.Abs(tempPrjDistanceValue) > 0 && Math.Round(Math.Abs(minPrjDistanceValue) - (Math.Abs(tempPrjDistanceValue)), 2) >= 0);

                            if (checkValue && checkDistanceXY)
                            {
                                maxIntersectValue = tempIntersectValue;
                                minPrjDistanceValue = tempPrjDistanceValue;
                                resultLSSolid = levelAndGridSolid;
                            }

                        }
                        #endregion

                        #region Првоеряю положение в секции по вектору Z
                        resSolid = GetNearestVerticalIntesectedInstSolid(instGeomData.IGDSolid, levelAndGridSolid);
                        if (resSolid != null && resSolid.Volume > 0)
                        {
                            XYZ resSolidCentroid = resSolid.ComputeCentroid();
                            if (resSolidCentroid == null)
                                continue;

                            foreach (Face levelAndGridFace in levelAndGridFaceArray)
                            {
                                IntersectionResult prjPointResult = levelAndGridFace.Project(resSolidCentroid);
                                if (prjPointResult == null)
                                    continue;

                                bool checkPntInside = levelAndGridFace.IsInside(prjPointResult.UVPoint);
                                bool checkDistanceZ = prjPointResult.Distance < minPrjDistanceValue;
                                if (checkPntInside && checkDistanceZ)
                                {
                                    minPrjDistanceValue = Math.Round(prjPointResult.Distance, 3);
                                    resultLSSolid = levelAndGridSolid;
                                }
                            }
                        }
                        #endregion

                    }
                    // Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                    catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                    {
                        if (!ErrorElements.Any(e => e.ErrorElement.Id == instGeomData.IEDElem.Id))
                        {
                            ErrorElements.Add(new GripParamError(
                                instGeomData.IEDElem,
                                "Геометрия: Элемент нужно назначить вручную (геометрию невозможно проанализиовать)"));
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!ErrorElements.Any(e => e.ErrorElement.Id == instGeomData.IEDElem.Id))
                        {
                            throw new GripParamExection($"Что-то непонятное с элементом с id: {instGeomData.IEDElem.Id}. Отправь разработчику:\n {ex.Message}");
                        }
                    }
                }
            }

            return resultLSSolid;
        }
    }
}