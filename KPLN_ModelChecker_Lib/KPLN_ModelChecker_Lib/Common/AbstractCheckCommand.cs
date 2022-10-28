using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_Lib.Common
{
    public abstract class AbstractCheckCommand
    {
        /// <summary>
        /// UIApplication для документа
        /// </summary>
        private protected UIApplication UIApp { get; set; }

        /// <summary>
        /// Список сущностей со своими атрибутами, которые соответсвуют ошибке
        /// </summary>
        private protected IList<ElementEntity> ErrorCollection { get; set; }

        /// <summary>
        /// Команда к выполнению
        /// </summary>
        public abstract IList<ElementEntity> Run();
    }
}
