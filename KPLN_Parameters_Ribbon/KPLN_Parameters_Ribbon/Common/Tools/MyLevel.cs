using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.Tools
{
    internal class MyLevel
    {

        private Level _currentLevel;
        private Level _aboveLevel;

        /// <summary>
        /// Текущий уровень
        /// </summary>
        public Level CurrentLevel
        {
            get
            {
                return _currentLevel;
            }
            private set
            {
                _currentLevel = value;
            }
        }

        /// <summary>
        /// Уровень выше
        /// </summary>
        public Level AboveLevel
        {
            get
            {
                return _aboveLevel;
            }
            private set
            {
                _aboveLevel = value;
            }
        }

        public MyLevel(Level currentLevel, Level aboveLevel)
        {
            _currentLevel = currentLevel;
            _aboveLevel = aboveLevel;
        }
    }
}
