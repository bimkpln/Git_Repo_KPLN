using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Обертка для элементов, используется при выявлении ошибок в работе скриптов
    /// </summary>
    internal class CheckCommandError
    {
        public CheckCommandError(Element errorElement, string errorMessage)
        {
            ErrorElement = errorElement;
            ErrorMessage = errorMessage;
        }

        public Element ErrorElement { get; }

        public string ErrorMessage { get; }
    }
}
