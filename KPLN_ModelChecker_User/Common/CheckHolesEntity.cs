using System;
using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI.HtmlWindow;

namespace KPLN_ModelChecker_User.Common
{
    /// <summary>
    /// Контейнер для сущностей плагина CommandCheckHoles
    /// </summary>
    internal class CheckHolesEntity
    {
        private Transform _currentLinkTransform;

        public CheckHolesEntity(Element elem)
        {
            CurrentElement = elem;
        }

        public CheckHolesEntity(Element elem, RevitLinkInstance linkInstance) : this(elem)
        {
            CurrentLinkInstance = linkInstance;
        }

        public Element CurrentElement { get; }

        public Solid CurrentSolid { get; set; }

        public BoundingBoxXYZ CurrentBBox { get; protected set; }

        public RevitLinkInstance CurrentLinkInstance { get; } = null;

        public Transform CurrentLinkTransform
        {
            get
            {
                if (_currentLinkTransform == null && CurrentLinkInstance != null)
                    _currentLinkTransform = CurrentLinkInstance.GetTotalTransform();
                return _currentLinkTransform;
            }
        }

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

                if (resultSolid != null)
                    CurrentSolid = CurrentLinkInstance == null ? resultSolid : SolidUtils.CreateTransformed(resultSolid, CurrentLinkTransform);
                // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК а также общих вложенных, которые часто также без геометрии.
                else if (CurrentElement is FamilyInstance famInst)
                {
                    string elem_type_param = famInst.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.ToLower() ?? "";
                    string elem_family_param = famInst.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString()?.ToLower() ?? "";

                    if (famInst.SuperComponent == null
                        && (
                            // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК
                            !elem_family_param.Contains("ostec") 
                            && !elem_family_param.Contains("470_dkc_s5_accessories")
                            && !elem_family_param.Contains("dkc_ceiling")
                            && !elem_family_param.Contains("757_опора_неподвижная")
                            // Фильтрация семейств под которое НИКОГДА не должно быть отверстий
                            && !elem_family_param.StartsWith("501_")
                            && !elem_family_param.StartsWith("551_")
                            && !elem_family_param.StartsWith("556_")
                            && !elem_family_param.StartsWith("557_")
                            && !elem_family_param.StartsWith("560_")
                            && !elem_family_param.StartsWith("561_")
                            && !elem_family_param.StartsWith("565_")
                            && !elem_family_param.StartsWith("570_")
                            && !elem_family_param.StartsWith("582_")
                            && !elem_family_param.StartsWith("592_")
                            // Фильтрация типов семейств для которых опытным путём определено, что солид у них не взять (очень сложные семейства)
                            && !elem_type_param.Contains("узел учета квартиры для гвс")
                            && !elem_type_param.Contains("узел учета офиса для гвс")
                            && !elem_type_param.Contains("узел учета квартиры для хвс")
                            && !elem_type_param.Contains("узел учета офиса для хвс")))
                            HtmlOutput.Print($"У элемента семейства {famInst.Symbol.FamilyName} " +
                            $"из модели {CurrentLinkInstance.Name} с id: {CurrentElement.Id} проблемы с получением Solid);", MessageType.Warning);
                }
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
