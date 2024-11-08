using Autodesk.Revit.DB;
using KPLN_Parameters_Ribbon.Forms;
using System.Collections.Generic;

namespace KPLN_Parameters_Ribbon.Common.CheckParam.Builder
{
    internal abstract class AbstrAuditBuilder
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
        /// Коллекция элементов для проверки
        /// </summary>
        private List<Element> _elemsToCheck = new List<Element>();
        /// <summary>
        /// Количество всех элементов
        /// </summary>
        private int _allElementsCount = 0;

        public AbstrAuditBuilder(Document doc, string docMainTitle)
        {
            Doc = doc;
            DocMainTitle = docMainTitle;
        }

        /// <summary>
        /// Коллекция элементов для проверки
        /// </summary>
        public List<Element> ElemsToCheck
        {
            get { return _elemsToCheck; }
            protected set { _elemsToCheck = value; }
        }

        /// <summary>
        /// Количество всех элементов
        /// </summary>
        public int AllElementsCount
        {
            get
            {
                if (_allElementsCount == 0)
                    _allElementsCount = ElemsToCheck.Count;

                return _allElementsCount;
            }
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

        }

        /// <summary>
        /// Метод проверки факта заполнения парамтеров
        /// </summary>
        /// <param name="pb">Прогресс-бар для визуализации процесса выполнения</param>
        /// <returns></returns>
        public abstract bool ExecuteParamsAudit(Progress_Single pb);
    }
}
