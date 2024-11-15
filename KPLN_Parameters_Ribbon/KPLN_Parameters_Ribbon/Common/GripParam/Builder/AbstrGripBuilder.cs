using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib.LevelAndGridBoxUtil;
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
        private List<LevelAndGridSolid> _sectDataSolids = new List<LevelAndGridSolid>();
        private List<GripParamError> _errorElements = new List<GripParamError>();
        private int _allElementsCount = 0;
        private int _hostElementsCount = 0;

        public AbstrGripBuilder(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, double floorScreedHeight, double downAndTopExtra)
        {
            Doc = doc;
            DocMainTitle = docMainTitle;
            LevelParamName = levelParamName;
            LevelNumberIndex = levelNumberIndex;
            SectionParamName = sectionParamName;
            FloorScreedHeight = floorScreedHeight;
            DownAndTopExtra = downAndTopExtra;
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
        /// Индекс, указывающий номер этажа, после разделения имени уровня по разделителю
        /// </summary>
        internal int LevelNumberIndex { get; }

        /// <summary>
        /// Имя параметра, в который осуществляется запись секции
        /// </summary>
        internal string SectionParamName { get; }

        /// <summary>
        /// Толщина смещения относительно уровня (чаще всего - стяжка пола). Нужна для перекидки значения элементов в стяжке на этаж выше
        /// </summary>
        internal double FloorScreedHeight { get; }

        /// <summary>
        /// Размер увеличения нижнего и вехнего боксов. Нужна для привязки элементов, расположенных за пределами крайних уровней
        /// </summary>
        internal double DownAndTopExtra { get; }

        /// <summary>
        /// Счетчик выпроненных операций по записи данных
        /// </summary>
        internal int PbCounter { get; private set; } = 0;

        /// <summary>
        /// Коллекция элементов на уровне
        /// </summary>
        public List<InstanceElemData> ElemsOnLevel
        {
            get { return _elemsOnLevel; }
            protected set { _elemsOnLevel = value; }
        }

        /// <summary>
        /// Коллекция элементов, которые являются общими вложенными
        /// </summary>
        public List<InstanceElemData> ElemsByHost
        {
            get { return _elemsByHost; }
            protected set { _elemsByHost = value; }
        }

        /// <summary>
        /// Коллекция элементов под уровне
        /// </summary>
        public List<InstanceElemData> ElemsUnderLevel
        {
            get { return _elemsUnderLevel; }
            protected set { _elemsUnderLevel = value; }
        }

        /// <summary>
        /// Коллекция лестниц
        /// </summary>
        public List<InstanceElemData> StairsElems
        {
            get { return _stairsElems; }
            protected set { _stairsElems = value; }
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
        public List<LevelAndGridSolid> SectDataSolids
        {
            get { return _sectDataSolids; }
            protected set { _sectDataSolids = value; }
        }


        /// <summary>
        /// Коллекция элементов, которые при анализе выдали ошибку
        /// </summary>
        public List<GripParamError> ErrorElements 
        { 
            get => _errorElements;
            private set
            {
                _errorElements = value;
            }
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

            ICollection<ElementId> availableWSElemsId = WorksharingUtils.CheckoutElements(doc, AllElements.Select(e => e.CurrentElem.Id).ToArray());
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
                InstanceGeomData instGeomData = (InstanceGeomData)instElemData 
                    ?? throw new GripParamExection(
                        $"Элемент {instElemData.CurrentElem.Id} был не правильно назначен (как элемент без гометриии. Обратись к разработчику\n");

                LevelAndGridSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.CurrentElem,
                            "Геометрия: Элементу не удалось присвоить данные по геометрии"));
                        continue;
                    }
                }

                // Кастомная настройка записи данных для пректа СЕТУНЬ
                if (Doc.Title.Contains("СЕТ_1"))
                {
                    string tempLvlData = maxIntersectInstance.CurrentLevelData.CurrentLevel.LookupParameter(LevelParamName).AsString().ToLower();
                    if (tempLvlData.Contains("кровля"))
                        instElemData.CurrentElem.LookupParameter(LevelParamName).Set("Кровля");
                    else
                        instElemData.CurrentElem.LookupParameter(LevelParamName).Set($"{maxIntersectInstance.CurrentLevelData.CurrentLevelNumber}_этаж");

                    string tempSectData = maxIntersectInstance.CurrentLevelData.CurrentSectionNumber;
                    if (tempSectData.Contains("С1"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 1");
                    else if (tempSectData.Contains("С2"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 2");
                    else if (tempSectData.Contains("С3"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 3");
                    else if (tempSectData.Contains("С4"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 4");
                    else if (tempSectData.Contains("СТЛ"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Паркинг");
                }
                else
                {
                    instElemData.CurrentElem.LookupParameter(LevelParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentLevelNumber);
                    instElemData.CurrentElem.LookupParameter(SectionParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentSectionNumber);
                }
                
                instElemData.IsEmptyData = false;

                pb.Update(++PbCounter, "Поиск по геометрии");
            }

            foreach (InstanceElemData instElemData in ElemsUnderLevel)
            {
                InstanceGeomData instGeomData = (InstanceGeomData)instElemData ??
                    throw new GripParamExection(
                        $"Элемент {instElemData.CurrentElem.Id} был не правильно назначен (как элемент без гометриии. Обратись к разработчику");

                LevelAndGridSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.CurrentElem,
                            "Геометрия: Элементу не удалось присвоить данные по геометрии"));
                        continue;
                    }
                }
                LevelAndGridSolid downLevelAndGridSolid = SectDataSolids
                    .Where(s =>
                        s.GridData.CurrentSection.Equals(maxIntersectInstance.GridData.CurrentSection)
                        && s.CurrentLevelData.CurrentLevel.Equals(maxIntersectInstance.CurrentLevelData.CurrentLevel))
                    .FirstOrDefault();

                instElemData.CurrentElem.LookupParameter(SectionParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentSectionNumber);

                // Кастомная настройка записи данных для пректа СЕТУНЬ
                if (Doc.Title.Contains("СЕТ_1"))
                {
                    string tempLvlData = downLevelAndGridSolid.CurrentLevelData.CurrentLevel.LookupParameter(LevelParamName).AsString().ToLower();
                    if (tempLvlData.Contains("кровля"))
                        instElemData.CurrentElem.LookupParameter(LevelParamName).Set("Кровля");
                    else
                        instElemData.CurrentElem.LookupParameter(LevelParamName).Set($"{downLevelAndGridSolid.CurrentLevelData.CurrentLevelNumber}_этаж");

                    string tempSectData = downLevelAndGridSolid.CurrentLevelData.CurrentSectionNumber;
                    if (tempSectData.Contains("С1"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 1");
                    else if (tempSectData.Contains("С2"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 2");
                    else if (tempSectData.Contains("С3"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 3");
                    else if (tempSectData.Contains("С4"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Секция 4");
                    else if (tempSectData.Contains("СТЛ"))
                        instElemData.CurrentElem.LookupParameter(SectionParamName).Set("Паркинг");
                }
                else
                {
                    instElemData.CurrentElem.LookupParameter(LevelParamName).Set(downLevelAndGridSolid.CurrentLevelData.CurrentLevelNumber);
                    instElemData.IsEmptyData = false;
                }
                

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
            foreach (InstanceElemData instElemData in ElemsByHost)
            {
                Element elem = instElemData.CurrentElem;
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
                        instElemData.CurrentElem,
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
            ErrorElements.AddRange(AllElements.Where(e => e.IsEmptyData).Select(e => new GripParamError(e.CurrentElem, "Элементы не подверглись анализу")));
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
                Element elem = instData.CurrentElem;
                Parameter sectParam = elem.LookupParameter(SectionParamName);
                Parameter levParam = elem.LookupParameter(LevelParamName);
                if (sectParam == null || levParam == null)
                {
                    throw new GripParamExection(
                        $"Прервано по причине отсутствия необходимых параметров захваток (секции или этажа). " +
                        $"Пример: \nКатегория: {elem.Category.Name} / id: {elem.Id}");
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
        private Solid GetNearestHorizontalIntesectedInstSolid(Solid instSolid, LevelAndGridSolid sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.CurrentSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(sectTransform.Origin.X, sectTransform.Origin.Y, instTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.CurrentSolid, BooleanOperationsType.Intersect);
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
        private Solid GetNearestVerticalIntesectedInstSolid(Solid instSolid, LevelAndGridSolid sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.CurrentSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(instTransform.Origin.X, instTransform.Origin.Y, sectTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.CurrentSolid, BooleanOperationsType.Intersect);
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
                    instElemData.CurrentElem,
                    $"Вложенность: У элемента основы (id: {hostElem.Id}) не заполнены данные для передачи"));
            }
            else
            {
                Parameter elemSectParam = instElemData.CurrentElem.LookupParameter(SectionParamName);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemSectParam.IsReadOnly)
                    elemSectParam.Set(hostElemSectParamValue);

                Parameter elemLevParam = instElemData.CurrentElem.LookupParameter(LevelParamName);
                // Вложенные семейства могут быть заблочены через формулу, для передачи из родительского
                if (!elemLevParam.IsReadOnly)
                    elemLevParam.Set(hostElemLevParamValue);
            }
        }

        /// <summary>
        /// Получить LevelAndGridSolid с наибольшим пересечением по солидам
        /// </summary>
        /// <param name="instGeomData">Элемент для анализа</param>
        private LevelAndGridSolid GetMaxIntersectedLevelAndGridSolid(InstanceGeomData instGeomData)
        {
            LevelAndGridSolid result = null;
            double maxIntersectValue = 0;
            foreach (LevelAndGridSolid levelAndGridSolid in SectDataSolids)
            {
                // Игнорирую заведомо отличающиеся по отметкам секции (9-10 м)
                if (Math.Abs(instGeomData.MinAndMaxElevation[0] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[0]) > 30
                    && Math.Abs(instGeomData.MinAndMaxElevation[1] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[1]) > 30)
                    continue;
                
                double tempIntersectValue = 0;
                try
                {
                    foreach (Solid instSolid in instGeomData.CurrentSolidColl)
                    {
                        if (instSolid.Volume == 0)
                            continue;

                        // Проверяю положение в секции
                        Solid checkIntersectSectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            instSolid, 
                            levelAndGridSolid.CurrentSolid, 
                            BooleanOperationsType.Intersect);

                        if (checkIntersectSectSolid == null || !(checkIntersectSectSolid.Volume > 0)) 
                            continue;
                        
                        tempIntersectValue += Math.Round(checkIntersectSectSolid.Volume, 10);
                    }

                    if (tempIntersectValue > 0 && maxIntersectValue < tempIntersectValue)
                    {
                        maxIntersectValue = tempIntersectValue;
                        result = levelAndGridSolid;
                    }
                }
                // Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    ErrorElements.Add(new GripParamError(
                        instGeomData.CurrentElem,
                        "Геометрия: Элемент нужно назначить вручную (геометрию невозможно проанализиовать)"));
                }
                catch (Exception ex)
                {
                    throw new GripParamExection($"Что-то непонятное с элементом с id: {instGeomData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}");
                }
            }

            return result;
        }

        /// <summary>
        /// Получить LevelAndGridSolid, который расположен ближе всего, и который имеет макс. объем по пересечению
        /// </summary>
        /// <param name="instGeomData">Элемент для анализа</param>
        private LevelAndGridSolid GetNearestIntersectedLevelAndGridSolid(InstanceGeomData instGeomData)
        {
            // Переменные для поиска ближайшей секции в плоскости XY
            LevelAndGridSolid resultHorizontal = null;
            double maxIntersectValue = double.MinValue;
            // 4,5 м - условно возможное отклонение элемента от солида
            double minPrjDistanceValue = 15;

            // Переменные для поиска ближайшей секции по вектору Z
            LevelAndGridSolid resultVertical = null;
            double minVerticalDistance = double.MinValue;
            // 4,5 м - условно возможное отклонение элемента от солида
            double minVerticalPrjDistanceValue = 15;

            foreach (LevelAndGridSolid levelAndGridSolid in SectDataSolids)
            {
                // Игнорирую заведомо отличающиеся по отметкам секции (9-10 м)
                if (Math.Abs(instGeomData.MinAndMaxElevation[0] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[0]) > 30
                    && Math.Abs(instGeomData.MinAndMaxElevation[1] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[1]) > 30)
                    continue;

                double tempIntersectValue = 0;
                double tempPrjDistanceValue = 15;
                try
                {
                    FaceArray levelAndGridFaceArray = levelAndGridSolid.CurrentSolid.Faces;

                    #region Првоеряю положение в секции в плоскости XY
                    foreach (Solid instSolid in instGeomData.CurrentSolidColl)
                    {
                        if (instSolid.Volume == 0)
                            continue;

                        Solid resSolid = GetNearestHorizontalIntesectedInstSolid(instSolid, levelAndGridSolid);
                        if (resSolid != null && resSolid.Volume > 0)
                            tempIntersectValue += Math.Round(resSolid.Volume, 3);
                    }

                    // Проверяю расстояние до секции с целью выявления ближайшей
                    if (tempIntersectValue > 0)
                    {
                        foreach (Face levelAndGridFace in levelAndGridFaceArray)
                        {
                            foreach (XYZ checkPoint in instGeomData.CurrentGeomCenterColl)
                            {
                                IntersectionResult prjPointResult = levelAndGridFace.Project(checkPoint);
                                if (prjPointResult != null && prjPointResult.Distance < tempPrjDistanceValue)
                                {
                                    tempPrjDistanceValue = Math.Round(prjPointResult.Distance, 3);
                                }
                            }
                        }
                    }

                    if ((tempIntersectValue > 0 && maxIntersectValue <= tempIntersectValue) && (tempPrjDistanceValue < 15 && minPrjDistanceValue >= tempPrjDistanceValue))
                    {
                        maxIntersectValue = tempIntersectValue;
                        minPrjDistanceValue = tempPrjDistanceValue;
                        resultHorizontal = levelAndGridSolid;
                    }
                    #endregion

                    #region Првоеряю положение в секции по вектору Z
                    foreach (Solid instSolid in instGeomData.CurrentSolidColl)
                    {
                        if (instSolid.Volume == 0)
                            continue;

                        Solid resSolid = GetNearestVerticalIntesectedInstSolid(instSolid, levelAndGridSolid);
                        if (resSolid != null && resSolid.Volume > 0)
                        {
                            foreach (Face levelAndGridFace in levelAndGridFaceArray)
                            {
                                foreach (XYZ checkPoint in instGeomData.CurrentGeomCenterColl)
                                {
                                    IntersectionResult prjPointResult = levelAndGridFace.Project(checkPoint);
                                    if (prjPointResult != null && levelAndGridFace.IsInside(prjPointResult.UVPoint) && prjPointResult.Distance < minVerticalPrjDistanceValue)
                                    {
                                        minVerticalPrjDistanceValue = Math.Round(prjPointResult.Distance, 3);
                                        resultVertical = levelAndGridSolid;
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                }
                // Отлов ошибки для сложной геометрии, для которой невозможно выполнить анализ на коллизии (нужно перемоделить элемент, что не приемлемо)
                catch (Autodesk.Revit.Exceptions.InvalidOperationException)
                {
                    if (!ErrorElements.Any(e => e.ErrorElement.Id == instGeomData.CurrentElem.Id))
                    {
                        ErrorElements.Add(new GripParamError(
                            instGeomData.CurrentElem,
                            "Геометрия: Элемент нужно назначить вручную (геометрию невозможно проанализиовать)"));
                    }
                }
                catch (Exception ex)
                {
                    if (!ErrorElements.Any(e => e.ErrorElement.Id == instGeomData.CurrentElem.Id))
                    {
                        throw new GripParamExection($"Что-то непонятное с элементом с id: {instGeomData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}");
                    }
                }
            }

            return resultHorizontal ?? resultVertical;
        }
    }
}