using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Обертка для ошибок, выявленных при работе скриптов
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
