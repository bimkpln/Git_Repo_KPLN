using Autodesk.Revit.DB;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Контейнер для хранения данных по проверяемым элементам Revit
    /// </summary>
    internal class InstanceElemData
    {
        public InstanceElemData(Element elem)
        {
            IEDElem = elem;
        }

        /// <summary>
        /// Элемент ревит
        /// </summary>
        public Element IEDElem { get; }

        /// <summary>
        /// Метка, указывающая статус анализа элемента. По умолчанию - true, т.е. данные не заполнялись
        /// </summary>
        public bool IsEmptyData { get; set; } = true;
    }
}
