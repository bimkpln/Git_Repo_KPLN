using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

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

        /// <summary>
        /// Элементы над уровнем
        /// </summary>
        protected private List<Element> _elemsOnLevel = new List<Element>();

        /// <summary>
        /// Коллекция всех элементов
        /// </summary>
        protected private List<Element> _allElems = new List<Element>();

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
        /// Метод заполнения парамтеров уровня
        /// </summary>
        public abstract bool ExecuteLevelParams();

        /// <summary>
        /// Метод заполнения парамтеров секции
        /// </summary>
        public abstract bool ExecuteSectionParams();
    }
}
