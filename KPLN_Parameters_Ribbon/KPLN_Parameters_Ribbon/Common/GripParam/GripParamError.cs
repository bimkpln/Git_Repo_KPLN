using Autodesk.Revit.DB;

namespace KPLN_Parameters_Ribbon.Common.GripParam
{
    /// <summary>
    /// Обертка для элементов, используется при выявлении ошибок в работе скриптов
    /// </summary>
    internal class GripParamError
    {
        public GripParamError(Element errorElement, string errorMessage)
        {
            ErrorElement = errorElement;
            ErrorMessage = errorMessage;
        }

        public Element ErrorElement { get; }

        public string ErrorMessage { get; }
    }
}
