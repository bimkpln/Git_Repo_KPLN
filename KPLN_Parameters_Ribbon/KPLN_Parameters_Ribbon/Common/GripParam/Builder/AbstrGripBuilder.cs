﻿using Autodesk.Revit.DB;
using KPLN_ModelChecker_Lib;
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
        /// Документ Ревит
        /// </summary>
        protected internal readonly Document Doc;

        /// <summary>
        /// Имя документа Ревит
        /// </summary>
        internal readonly string DocMainTitle;

        /// <summary>
        /// Имя параметра, в который осуществляется запись уровня
        /// </summary>
        internal readonly string LevelParamName;

        /// <summary>
        /// Индекс, указывающий номер этажа, после разделения имени уровня по разделителю
        /// </summary>
        internal readonly int LevelNumberIndex;

        /// <summary>
        /// Имя параметра, в который осуществляется запись секции
        /// </summary>
        internal readonly string SectionParamName;

        /// <summary>
        /// Толщина смещения относительно уровня (чаще всего - стяжка пола). Нужна для перекидки значения элементов в стяжке на этаж выше
        /// </summary>
        internal readonly double FloorScreedHeight;

        /// <summary>
        /// Размер увеличения нижнего и вехнего боксов. Нужна для привязки элементов, расположенных за пределами крайних уровней
        /// </summary>
        internal readonly double DownAndTopExtra;

        /// <summary>
        /// Счетчик выпроненных операций по записи данных
        /// </summary>
        internal int PbCounter = 0;

        /// <summary>
        ///  GUID параметра для исключения перезаписи ("ПЗ_Перезаписать")
        /// </summary>
        private readonly Guid _revalueParamGuid = new Guid("38157d2d-f952-41e5-8d05-2c962addfe56");
        private List<InstanceElemData> _elemsOnLevel = new List<InstanceElemData>();
        private List<InstanceElemData> _elemsByHost = new List<InstanceElemData>();
        private List<InstanceElemData> _elemsUnderLevel = new List<InstanceElemData>();
        private List<InstanceElemData> _stairsElems = new List<InstanceElemData>();
        private List<LevelAndGridSolid> _sectDataSolids = new List<LevelAndGridSolid>();
        private int _allElementsCount = 0;
        private int _hostElementsCount = 0;

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
        /// Количество всех элементов
        /// </summary>
        public int AllElementsCount
        {
            get { return _allElementsCount; }
            private set { _allElementsCount = value; }
        }

        /// <summary>
        /// Количество элементов с основанием
        /// </summary>
        public int HostElementsCount
        {
            get { return _hostElementsCount; }
            private set { _hostElementsCount = value; }
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
        public List<GripParamError> ErrorElements = new List<GripParamError>();

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
        /// Метод подготовки элементов к обработке
        /// </summary>
        public abstract void Prepare();

        /// <summary>
        /// Метод проверки элементов на заполняемость параметров и чистка списков от элементов, не подверженных чистке
        /// </summary>
        public virtual void Check()
        {
            Task elemsOnLevelCheckTask = Task.Run(() => CheckElemParams(ElemsOnLevel));
            Task elemsByHostCheckTask = Task.Run(() => CheckElemParams(ElemsByHost));
            Task elemsUnderLevelCheckTask = Task.Run(() => CheckElemParams(ElemsUnderLevel));
            Task elemsStairsElemsCheckTask = Task.Run(() => CheckElemParams(StairsElems));

            Task.WaitAll(new Task[] { elemsOnLevelCheckTask, elemsByHostCheckTask, elemsUnderLevelCheckTask, elemsStairsElemsCheckTask });
        }

        /// <summary>
        /// Подсчет элементов для обработки
        /// </summary>
        public virtual void CountElements()
        {
            HostElementsCount = ElemsByHost.Count;
            AllElementsCount = ElemsOnLevel.Count + ElemsUnderLevel.Count + ElemsByHost.Count + StairsElems.Count;
            if (AllElementsCount == 0)
                throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет, или имя проекта не соответсвует ВЕР!");
        }

        /// <summary>
        /// Метод заполнения парамтеров захваток (секции и уровня) по анализу геометрии
        /// </summary>
        /// <param name="pb">Прогресс-бар для визуализации процесса выполнения</param>
        public void ExecuteGripParams_ByGeom(Progress_Single pb)
        {
            foreach (InstanceElemData instElemData in ElemsOnLevel)
            {
                //if (instElemData.CurrentElem.Id.IntegerValue == 5137353)
                //{
                //    var a = 1;
                //}
                InstanceGeomData instGeomData = (InstanceGeomData)instElemData;
                if (instGeomData == null) throw new Exception($"Элемент {instElemData.CurrentElem.Id} был не правильно назначен (как элемент ьбез гометриии. Обратись к разработчику");

                LevelAndGridSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.CurrentElem,
                            "Геометрия: Элементу не удалось присвоить данные по геомертии"));
                        continue;
                    }
                }

                instElemData.CurrentElem.LookupParameter(SectionParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentSectionNumber);
                instElemData.CurrentElem.LookupParameter(LevelParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentLevelNumber);
                pb.Update(++PbCounter, "Поиск по геометрии");
            }

            foreach (InstanceElemData instElemData in ElemsUnderLevel)
            {
                //if (instElemData.CurrentElem.Id.IntegerValue == 5137353)
                //{
                //    var a = 1;
                //}
                InstanceGeomData instGeomData = (InstanceGeomData)instElemData;
                if (instGeomData == null) throw new Exception($"Элемент {instElemData.CurrentElem.Id} был не правильно назначен (как элемент ьбез гометриии. Обратись к разработчику");

                LevelAndGridSolid maxIntersectInstance = GetMaxIntersectedLevelAndGridSolid(instGeomData);
                if (maxIntersectInstance == null)
                {
                    // Повторная проходка для элементов, которые находятся ВНЕ секции
                    maxIntersectInstance = GetNearestIntersectedLevelAndGridSolid(instGeomData);
                    if (maxIntersectInstance == null)
                    {
                        ErrorElements.Add(new GripParamError(
                            instElemData.CurrentElem,
                            "Геометрия: Элементу не удалось присвоить данные по геомертии"));
                        continue;
                    }
                }
                LevelAndGridSolid downLevelAndGridSolid = SectDataSolids
                    .Where(s =>
                        s.GridData.CurrentSection.Equals(maxIntersectInstance.GridData.CurrentSection)
                        && s.CurrentLevelData.CurrentLevel.Equals(maxIntersectInstance.CurrentLevelData.CurrentLevel))
                    .FirstOrDefault();

                instElemData.CurrentElem.LookupParameter(SectionParamName).Set(maxIntersectInstance.CurrentLevelData.CurrentSectionNumber);
                instElemData.CurrentElem.LookupParameter(LevelParamName).Set(downLevelAndGridSolid.CurrentLevelData.CurrentLevelNumber);
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
            if (ElemsByHost.Count == 0)
            {
                pb.Update(PbCounter, "Анализ элементов на основе");
            }
            else
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

                    pb.Update(++PbCounter, "Анализ элементов на основе");
                }
            }
        }

        /// <summary>
        /// Метод проверки и очичтки элементов
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
                    throw new Exception($"Прервано по причине отсутствия необходимых параметров захваток (секции или этажа). " +
                        $"Пример: \nКатегория: {elem.Category.Name} / id: {elem.Id}");
                }

                Parameter canReValueParam = elem.get_Parameter(_revalueParamGuid);
                return canReValueParam != null && canReValueParam.HasValue && canReValueParam.AsInteger() != 1;
            });
        }

        /// <summary>
        /// Определение солида, который пересекается с солидом секции. Солид эл-та Ревит ПРИТЯГИВАЕТСЯ к солиду секции
        /// </summary>
        /// <param name="instSolid">Солид эл-та ревит для проверки</param>
        /// <param name="sectData">Солид секции для проверки</param>
        /// <returns></returns>
        private Solid GetIntesectedInstSolid(Solid instSolid, LevelAndGridSolid sectData)
        {
            // Необходимо "притянуть" через Transform элемент в центр солида секции, чтобы улучшить точность подсчета
            Transform sectTransform = sectData.LevelSolid.GetBoundingBox().Transform;
            Transform instTransform = instSolid.GetBoundingBox().Transform;
            Transform instInverseTransform = instTransform.Inverse;
            Solid instZerotransformSolid = SolidUtils.CreateTransformed(instSolid, instInverseTransform);
            sectTransform.Origin = new XYZ(sectTransform.Origin.X, sectTransform.Origin.Y, instTransform.Origin.Z);

            Solid transformedBySectdInstSolid = SolidUtils.CreateTransformed(instZerotransformSolid, sectTransform);
            Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(transformedBySectdInstSolid, sectData.LevelSolid, BooleanOperationsType.Intersect);
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
                // Игнорирую заведомо отличающиеся по отметкам секции
                if (Math.Abs(instGeomData.MinAndMaxElevation[0] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[0]) > 10
                    && Math.Abs(instGeomData.MinAndMaxElevation[1] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[1]) > 10)
                    continue;

                double tempIntersectValue = 0;
                try
                {
                    foreach (Solid instSolid in instGeomData.CurrentSolidColl)
                    {
                        if (instSolid.Volume == 0)
                            continue;

                        // Првоеряю положение в секции
                        Solid checkSectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(instSolid, levelAndGridSolid.LevelSolid, BooleanOperationsType.Intersect);
                        if (checkSectSolid != null && checkSectSolid.Volume > 0)
                        {
                            Solid resSolid = GetIntesectedInstSolid(instSolid, levelAndGridSolid);
                            if (resSolid != null)
                                tempIntersectValue += resSolid.Volume;
                        }
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
                    throw new Exception($"Что-то непонятное с элементом с id: {instGeomData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}");
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
            LevelAndGridSolid result = null;
            double maxIntersectValue = 0;
            foreach (LevelAndGridSolid levelAndGridSolid in SectDataSolids)
            {
                // Игнорирую заведомо отличающиеся по отметкам секции
                if (Math.Abs(instGeomData.MinAndMaxElevation[0] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[0]) > 10
                    && Math.Abs(instGeomData.MinAndMaxElevation[1] - levelAndGridSolid.CurrentLevelData.MinAndMaxLvlPnts[1]) > 10)
                    continue;

                double tempIntersectValue = 0;
                try
                {
                    foreach (Solid instSolid in instGeomData.CurrentSolidColl)
                    {
                        if (instSolid.Volume == 0)
                            continue;

                        // Првоеряю положение в секции
                        Solid resSolid = GetIntesectedInstSolid(instSolid, levelAndGridSolid);
                        if (resSolid != null && resSolid.Volume > 0)
                            tempIntersectValue += resSolid.Volume;
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
                    if (!ErrorElements.Where(e => e.ErrorElement.Id == instGeomData.CurrentElem.Id).Any())
                    {
                        ErrorElements.Add(new GripParamError(
                            instGeomData.CurrentElem,
                            "Геометрия: Элемент нужно назначить вручную (геометрию невозможно проанализиовать)"));
                    }
                }
                catch (Exception ex)
                {
                    if (!ErrorElements.Where(e => e.ErrorElement.Id == instGeomData.CurrentElem.Id).Any())
                    {
                        throw new Exception($"Что-то непонятное с элементом с id: {instGeomData.CurrentElem.Id}. Отправь разработчику:\n {ex.Message}");
                    }
                }
            }

            return result;
        }
    }
}
