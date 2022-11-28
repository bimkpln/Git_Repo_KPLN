using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Library_Forms.Common
{
    /// <summary>
    /// Коллекция результирующих статусов формы
    /// </summary>
    public class UIStatus
    {
        /// <summary>
        /// Статус окна после нажатия на кнопку Пуск
        /// </summary>
        public enum RunStatus { Run, Cancel, Close, CloseBecauseError }
    }
}
