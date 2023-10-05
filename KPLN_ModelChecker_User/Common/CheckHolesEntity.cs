using Autodesk.Revit.DB;
using System;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сущностей плагина CommandCheckHoles
    /// </summary>
    internal class CheckHolesEntity
    {
        public CheckHolesEntity(Element elem)
        {
            CurrentElement = elem;
        }
        
        public Element CurrentElement { get; }

        public Solid CurrentSolid { get; set; }

        public BoundingBoxXYZ CurrentBBox { get; protected set; }

        /// <summary>
        /// Заполнить поля RoomSolid и RoomBBox, если он не были заданы ранее (РЕСУРСОЕМКИЙ МЕТОД)
        /// </summary>
        /// <param name="detailLevel">Уровень детализации</param>
        public void SetGeometryData(ViewDetailLevel detailLevel)
        {
            #region Задаю Solid, если ранее не был создан
            if (CurrentSolid == null)
            {
                Solid resultSolid = null;
                GeometryElement geomElem = CurrentElement.get_Geometry(new Options { DetailLevel = detailLevel });
                foreach (GeometryObject gObj in geomElem)
                {
                    Solid solid = gObj as Solid;
                    GeometryInstance gInst = gObj as GeometryInstance;
                    if (solid != null) resultSolid = solid;
                    else if (gInst != null)
                    {
                        GeometryElement instGeomElem = gInst.GetInstanceGeometry();
                        double tempVolume = 0;
                        foreach (GeometryObject gObj2 in instGeomElem)
                        {
                            solid = gObj2 as Solid;
                            if (solid != null && solid.Volume > tempVolume)
                            {
                                tempVolume = solid.Volume;
                                resultSolid = solid;
                            }
                        }
                    }
                }
                CurrentSolid = resultSolid;
            }
            #endregion

            #region Задаю BoundingBoxXYZ, если ранее не был создан
            if (CurrentBBox == null)
            {
                if (CurrentSolid != null)
                {
                    BoundingBoxXYZ bbox = CurrentSolid.GetBoundingBox();
                    if (bbox == null)
                        throw new Exception($"Элементу {CurrentElement.Id} - невозможно создать BoundingBoxXYZ. Отправь сообщение разработчику");
                    Transform transform = bbox.Transform;

                    CurrentBBox = new BoundingBoxXYZ()
                    {
                        Max = transform.OfPoint(bbox.Max),
                        Min = transform.OfPoint(bbox.Min),
                    };
                }
            }
            #endregion
        }
    }
}
