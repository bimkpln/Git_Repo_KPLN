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
        /// Нужно ли использовать отдельные рабочие наборы для копируемых с мониторингом элементам?
        /// </summary>
        public bool UseMonitoredElements { get; set; }

        /// <summary>
        /// Имя рабоего набора для скопированных элементов (кроме осей и уровней)
        /// </summary>
        public string MonitoredElementsName { get; set; }

        /// <summary>
        /// Список параметров, для фильтрации
        /// </summary>
        public List<WorksetByCurrentParameter> WorksetByCurrentParameterList = new List<WorksetByCurrentParameter>();
    }
}
