using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сбора информации по каждому элементу ИОС
    /// </summary>
    internal class CheckHolesMEPData : CheckHolesEntity
    {
        public CheckHolesMEPData(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Коллеция точек элемента в пространстве
        /// </summary>
        public List<XYZ> CurrentLocationColl { get; set; } = new List<XYZ>(3);
    }
}
