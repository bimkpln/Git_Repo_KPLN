using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using KPLN_Parameters_Ribbon.Common.Tools;
using KPLN_Parameters_Ribbon.Forms;

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
            protected set { _allElementsCount = value; }
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
        public abstract bool Prepare();

        /// <summary>
        /// Метод проверки элементов на заполняемость параметров и чистка списков от элементов, не подверженных чистке
        /// </summary>
        public virtual bool Check()
        {
            return CheckElemParams(ElemsOnLevel)
            && CheckElemParams(ElemsByHost)
            && CheckElemParams(ElemsUnderLevel)
            && CheckElemParams(StairsElems);

        }

        /// <summary>
        /// Метод заполнения парамтеров уровня
        /// </summary>
        public abstract bool ExecuteGripParams(Progress_Single pb);

        /// <summary>
        /// Метод заполнения парамтеров уровня
        /// </summary>
        public abstract bool ExecuteLevelParams(Progress_Single pb);

        /// <summary>
        /// Метод заполнения парамтеров секции
        /// </summary>
        public abstract bool ExecuteSectionParams(Progress_Single pb);

        private bool CheckElemParams(List<Element> checkColl)
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
                    // Првоерка на возможность перезаписи по параметру "ПЗ_Перещаписать" и удаление нужных элементов 
                    if (canReValueParam != null && (canReValueParam.HasValue && canReValueParam.AsInteger() != 1))
                    {
                        checkColl.Remove(elem);
                    }
                }

                return true;
            }
            else
            {
                return true;
            }
        }
    }
}
