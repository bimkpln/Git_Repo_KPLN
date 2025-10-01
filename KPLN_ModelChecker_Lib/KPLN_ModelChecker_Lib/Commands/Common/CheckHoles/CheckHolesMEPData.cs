using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_ModelChecker_Lib.Commands.Common.CheckHoles
{
    /// <summary>
    /// Контейнер для сбора информации по каждому элементу ИОС
    /// </summary>
    internal sealed class CheckHolesMEPData : CheckHolesEntity
    {
        private readonly List<XYZ> _currentLocationColl = new List<XYZ>();

        public CheckHolesMEPData(Element elem, RevitLinkInstance linkInstance) : base(elem, linkInstance)
        {
        }

        /// <summary>
        /// Коллеция точек элемента в пространстве, с поправкой на координаты
        /// </summary>
        public List<XYZ> CurrentLocationColl
        {
            get => _currentLocationColl;
            set
            {
                foreach (XYZ xyz in value)
                {
                    XYZ prepearedXYZ = CurrentLinkInstance == null ? xyz : CurrentLinkTransform.OfPoint(xyz);
                    _currentLocationColl.Add(prepearedXYZ);
                }
            }
        }
    }
}
