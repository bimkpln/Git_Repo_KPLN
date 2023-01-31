using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Linq;

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
        /// Разделитель для уровня
        /// </summary>
        protected internal char SplitLevelChar = '_';

        private List<Element> _elemsOnLevel = new List<Element>();

        private List<Element> _elemsByHost = new List<Element>();

        private List<Element> _elemsInsulation = new List<Element>();

        private List<Element> _elemsUnderLevel = new List<Element>();

        private List<Element> _stairsElems = new List<Element>();

        private int _allElementsCount = 0;

        /// <summary>
        /// Коллекция элементов на уровне
        /// </summary>
        public List<Element> ElemsOnLevel
        {
            get { return _elemsOnLevel; }
            protected set { _elemsOnLevel = value; }
        }

        /// <summary>
        /// Коллекция элементов, которые являются общими вложенными
        /// </summary>
        public List<Element> ElemsByHost
        {
            get { return _elemsByHost; }
            protected set { _elemsByHost = value; }
        }

        /// <summary>
        /// Коллекция элементов изоляции
        /// </summary>
        public List<Element> ElemsInsulation
        {
            get { return _elemsInsulation; }
            protected set { _elemsInsulation = value; }
        }

        /// <summary>
        /// Коллекция элементов под уровне
        /// </summary>
        public List<Element> ElemsUnderLevel
        {
            get { return _elemsUnderLevel; }
            protected set { _elemsUnderLevel = value; }
        }

        /// <summary>
        /// Коллекция лестниц
        /// </summary>
        public List<Element> StairsElems
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

        public AbstrGripBuilder(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName)
        {
            Doc = doc;
            DocMainTitle = docMainTitle;
            LevelParamName = levelParamName;
            LevelNumberIndex = levelNumberIndex;
            SectionParamName = sectionParamName;
        }

        public AbstrGripBuilder(Document doc, string docMainTitle, string levelParamName, int levelNumberIndex, string sectionParamName, char splitLevelChar) : this(doc, docMainTitle, levelParamName, levelNumberIndex, sectionParamName)
        {
            SplitLevelChar = splitLevelChar;
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
            CheckElemParams(ElemsOnLevel);
            CheckElemParams(ElemsByHost);
            CheckElemParams(ElemsUnderLevel);
            CheckElemParams(StairsElems);
        }

        /// <summary>
        /// Подсчет элементов для обработки
        /// </summary>
        public virtual void CountElements()
        {
            AllElementsCount = ElemsOnLevel.Count + ElemsUnderLevel.Count + ElemsByHost.Count + ElemsInsulation.Count + StairsElems.Count;
            if (AllElementsCount == 0)
                throw new Exception("KPLN: Ошибка при взятии элементов из проекта. Таких категорий нет, или имя проекта не соответсвует ВЕР!");
        }

        /// <summary>
        /// Метод заполнения парамтеров захваток (секции и уровня)
        /// </summary>
        /// <param name="pb">Прогресс-бар для визуализации процесса выполнения</param>
        /// <returns></returns>
        public abstract bool ExecuteGripParams(Progress_Single pb);

        /// <summary>
        /// Метод проверки элементов
        /// </summary>
        /// <param name="checkColl">Коллекция для проверки</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private void CheckElemParams(List<Element> checkColl)
        {
            if (checkColl.Count > 0)
            {
                foreach (Element elem in checkColl.ToList())
                {
                    Parameter sectParam = elem.LookupParameter(SectionParamName);
                    Parameter levParam = elem.LookupParameter(LevelParamName);
                    if (sectParam == null || levParam == null)
                    {
                        throw new Exception($"Прервано по причине отсутсвия необходимых парамтеров захваток (секции или этажа). " +
                            $"Пример: Категория: {elem.Category.Name} / id: {elem.Id}");
                    }

                    Parameter canReValueParam = elem.get_Parameter(new Guid("38157d2d-f952-41e5-8d05-2c962addfe56"));
                    // Првоерка на возможность перезаписи по параметру "ПЗ_Перезаписать" и удаление нужных элементов 
                    if (canReValueParam != null && (canReValueParam.HasValue && canReValueParam.AsInteger() != 1))
                    {
                        checkColl.Remove(elem);
                    }
                }
            }
        }
    }
}
