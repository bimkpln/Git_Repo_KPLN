﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

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

        public CheckHolesEntity(Element elem, RevitLinkInstance linkInstance) : this (elem)
        {
            CurrentLinkInstance = linkInstance;
            CurrentLinkTransform = linkInstance.GetTotalTransform();
        }
        
        public Element CurrentElement { get; }

        public Solid CurrentSolid { get; set; }

        public BoundingBoxXYZ CurrentBBox { get; protected set; }
        
        public RevitLinkInstance CurrentLinkInstance { get; protected set; } = null;

        public Transform CurrentLinkTransform { get; protected set; } = null;

        /// <summary>
        /// Заполнить поля CurrentRoomSolid и CurrentRoomBBox, если он не были заданы ранее (РЕСУРСОЕМКИЙ МЕТОД)
        /// </summary>
        /// <param name="detailLevel">Уровень детализации</param>
        public void SetGeometryData(ViewDetailLevel detailLevel, List<CheckCommandError> notCriticalErrorElemColl)
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

                if (resultSolid != null)
                    CurrentSolid = CurrentLinkInstance == null ? resultSolid : SolidUtils.CreateTransformed(resultSolid, CurrentLinkTransform);
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
                    Transform resultTransform = CurrentLinkInstance == null ? transform : transform * CurrentLinkTransform;

                    CurrentBBox = new BoundingBoxXYZ()
                    {
                        Max = resultTransform.OfPoint(bbox.Max),
                        Min = resultTransform.OfPoint(bbox.Min),
                    };
                }
            }
            #endregion
        }
    }
}
