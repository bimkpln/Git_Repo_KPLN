using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_Lib
{
    /// <summary>
    /// Обертка для элементов, используется при выявлении ошибок в работе скриптов
    /// </summary>
    public sealed class CheckCommandError
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
