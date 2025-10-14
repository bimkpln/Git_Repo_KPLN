using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_Lib.Services.GripGeom.Core
{
    /// <summary>
    /// Контейнер для хранения данных по проверяемым элементам Revit
    /// </summary>
    public class InstanceElemData
    {
        public InstanceElemData(Element elem)
        {
            IEDElem = elem;
        }

        /// <summary>
        /// Элемент ревит
        /// </summary>
        public Element IEDElem { get; protected set; }

        /// <summary>
        /// Метка, указывающая статус анализа элемента. По умолчанию - true, т.е. данные не заполнялись
        /// </summary>
        public bool IsEmptyData { get; set; } = true;
    }
}
