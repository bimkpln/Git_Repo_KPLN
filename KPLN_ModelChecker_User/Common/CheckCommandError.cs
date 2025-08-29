using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Обертка для элементов, используется при выявлении ошибок в работе скриптов
    /// </summary>
    internal class CheckCommandErrorOld
    {
        public CheckCommandErrorOld(Element errorElement, string errorMessage)
        {
            ErrorElement = errorElement;
            ErrorMessage = errorMessage;
        }

        public Element ErrorElement { get; }

        public string ErrorMessage { get; }
    }
}
