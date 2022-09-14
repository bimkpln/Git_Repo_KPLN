using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    /// <summary>
    /// Контейнер для данных по солиду
    /// </summary>
    internal class MySolid
    {
        private Solid _solid;
        private string _levelIndex;
        private string _sectionIndex;

        /// <summary>
        /// Солид
        /// </summary>
        public Solid Solid 
        { 
            get 
            { 
                return _solid; 
            } 
            private set 
            { 
                _solid = value;
            } 
        }

        /// <summary>
        /// Индекс уровня
        /// </summary>
        public string LevelIndex
        {
            get
            {
                return _levelIndex;
            }
            private set
            {
                _levelIndex = value;
            }
        }

        /// <summary>
        /// Индекс секции
        /// </summary>
        public string SectionIndex
        {
            get
            {
                return _sectionIndex;
            }
            private set
            {
                _sectionIndex = value;
            }
        }

        public MySolid(Solid solid, string levelIndex, string sectionIndex)
        {
            _solid = solid;
            _levelIndex = levelIndex;
            _sectionIndex = sectionIndex;
        }

    }
}
