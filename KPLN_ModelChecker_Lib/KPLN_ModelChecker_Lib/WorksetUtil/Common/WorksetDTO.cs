using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.WorksetUtil.Common
{
    [Serializable]
    public class WorksetDTO
    {
        /// <summary>
        /// Префикс для связей
        /// </summary>
        public string LinkedFilesPrefix { get; set; }

        /// <summary>
        /// Имя рабоего набора для dwg-линков
        /// </summary>
        public string DWGLinksName { get; set; }

        /// <summary>
        /// Нужно показывать РН ГЛОБАЛЬНО для линков?
        /// </summary>
        public bool GlobalVisibleForLinks { get; set; }

        /// <summary>
        /// Список параметров, для фильтрации
        /// </summary>
        public List<WorksetByCurrentParameter> WorksetByCurrentParameterList = new List<WorksetByCurrentParameter>();
    }
}
