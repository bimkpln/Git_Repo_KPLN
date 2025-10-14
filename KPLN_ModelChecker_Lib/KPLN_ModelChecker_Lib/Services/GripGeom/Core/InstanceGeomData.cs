using Autodesk.Revit.DB;

namespace KPLN_ModelChecker_Lib.Services.GripGeom.Core
{
    /// <summary>
    /// Контейнер для хранения геометрических данных для InstanceElemData
    /// </summary>
    public class InstanceGeomData : InstanceElemData
    {
        private XYZ _igdGeomCenter;
        private BoundingBoxXYZ _igdBBox;
        private Solid _igdSolid;

        public InstanceGeomData(Element elem) : base(elem)
        {
        }

        /// <summary>
        /// Коллекция Solid элемента
        /// </summary>
        public Solid IGDSolid
        {
            get
            {
                if (_igdSolid == null)
                    _igdSolid = GeometryWorker.GetRevitElemUniontSolid(IEDElem);

                return _igdSolid;
            }

            protected set => _igdSolid = value;
        }

        /// <summary>
        /// Коллекция BoundingBoxXYZ элемента
        /// </summary>
        public BoundingBoxXYZ IGDBBox
        {
            get
            {
                if (_igdBBox == null)
                    _igdBBox = GeometryWorker.GetBoundingBoxXYZ(IGDSolid);

                return _igdBBox;
            }
        }

        /// <summary>
        /// Коллекция точек геом. центров составного элемента
        /// </summary>
        public XYZ IGDGeomCenter
        {
            get
            {
                if (_igdGeomCenter == null)
                {
                    try
                    {
                        _igdGeomCenter = IGDSolid.ComputeCentroid();
                    }
                    // Для сожной геометрии (разуклонка больших перекрытий) - могут быть проблемы с центроидом
                    catch
                    {
                        _igdGeomCenter = 0.5 * (IGDBBox.Max + IGDBBox.Min);
                    }
                }

                return _igdGeomCenter;
            }
        }
    }
}
