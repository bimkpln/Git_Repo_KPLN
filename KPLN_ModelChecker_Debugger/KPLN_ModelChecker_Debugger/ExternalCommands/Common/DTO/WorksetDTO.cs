using KPLN_ModelChecker_Debugger.ExternalCommands.Common.WorksetModels;
using System;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Debugger.ExternalCommands.Common
{
    [Serializable]
    public class WorksetDTO
    {
        /// <summary>
        /// Имя отдела
        /// </summary>
        public string Department;

        /// <summary>
        /// Префикс для связей
        /// </summary>
        public string LinkedFilesPrefix;

        /// <summary>
        /// Префикс для скопированных элементов
        /// </summary>
        public string CopyElementsPrefix;

        /// <summary>
        /// Список параметров, для фильтрации
        /// </summary>
        public List<WorksetByCurrentParameter> WorksetByCurrentParameterList = new List<WorksetByCurrentParameter>();
    }
}
