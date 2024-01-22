using System.Collections.Generic;

namespace KPLN_BIMTools_Ribbon.Common
{
    internal interface ICopier
    {
        /// <summary>
        /// Коллекция путей ОТКУДА копировать
        /// </summary>
        List<string> FilePathesFrom { get; set; }

        /// <summary>
        /// Коллекция путей КУДА копировать
        /// </summary>
        List<string> FilePathesTo { get; set; }

        /// <summary>
        /// Заполнить пути из файла-конфигурации
        /// </summary>
        /// <param name="pathToConfig">Путь к конфигу</param>
        void SetPathes(string pathToConfig);
    }
}
