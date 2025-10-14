using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib
{
    /// <summary>
    /// Обертка для элементов, используется для передачи информации пользователю
    /// </summary>
    public sealed class CheckCommandMsg
    {
        public CheckCommandMsg(Element msgElement, string message)
        {
            MsgElement = msgElement;
            Message = message;
        }

        /// <summary>
        /// Отдельный элемент для сообщения
        /// </summary>
        public Element MsgElement { get; }
        
        /// <summary>
        /// Сообщению пользователю
        /// </summary>
        public string Message { get; }
    }
}
