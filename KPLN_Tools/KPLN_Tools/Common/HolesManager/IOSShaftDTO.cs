using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Tools.Common.HolesManager
{
    /// <summary>
    /// Класс-контейнер для объединения данных по шахтам
    /// </summary>
    public class IOSShaftDTO
    {
        public FamilyInstance CurrentHole { get; set; }

        public Element DownFloorBinding { get; set; }
        
        /// <summary>
        /// Отметка нижнего элемента
        /// </summary>
        public double DownBindingElevation { get; set; }
        
        /// <summary>
        /// Приставка имени для отверстия
        /// </summary>
        public string BindingPrefixString { get; set; }

        /// <summary>
        /// Абсолютнгая отметка
        /// </summary>
        public double AbsElevation { get; set; }

        /// <summary>
        /// Относительная отметка
        /// </summary>
        public double RlvElevation { get; set; }

    }
}
