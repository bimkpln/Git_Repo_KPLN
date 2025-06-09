using Autodesk.Revit.DB;
using System.Collections.Generic;
using System;
using System.Linq;

namespace KPLN_OpeningHoleManager.Core
{
    /// <summary>
    /// Сущность эл-в ИОС для кэширования информации
    /// </summary>
    internal sealed class IOSElemEntity
    {
        private static LogicalOrFilter _elemCatLogicalOrFilter;
        private static Func<Element, bool> _elemFilterFunc;

        internal IOSElemEntity(Document linkDoc, Transform linkTrans, Element elem, Solid elemSolid, Solid aRIOS_IntesectionSolid)
        {
            IOS_LinkDocument = linkDoc;
            IOS_LinkTransform = linkTrans;
            IOS_Element = elem;
            IOS_Solid = elemSolid;
            ARKRIOS_IntesectionSolid = aRIOS_IntesectionSolid;
        }

        /// <summary>
        /// Ссылка на документ линка
        /// </summary>
        internal Document IOS_LinkDocument { get; private set; }

        /// <summary>
        /// Ссылка на Transform для линка
        /// </summary>
        internal Transform IOS_LinkTransform { get; private set; }

        /// <summary>
        /// Ссылка на элемент модели
        /// </summary>
        internal Element IOS_Element { get; private set; }

        /// <summary>
        /// Кэширование SOLID геометрии
        /// </summary>
        internal Solid IOS_Solid { get; private set; }

        /// <summary>
        /// Кэширование SOLID геометрии ПЕРЕСЕЧЕНИЯ между АР и ИОС
        /// </summary>
        internal Solid ARKRIOS_IntesectionSolid { get; private set; }

        /// <summary>
        /// Фильтр для ключевой фильтрации по категориям
        /// </summary>
        internal static LogicalOrFilter ElemCatLogicalOrFilter
        {
            get
            {
                if (_elemCatLogicalOrFilter == null)
                {
                    List<ElementFilter> catFilters = new List<ElementFilter>();
                    catFilters.AddRange(BuiltInCategories.Select(bic => new ElementCategoryFilter(bic)));

                    _elemCatLogicalOrFilter = new LogicalOrFilter(catFilters);
                }

                return _elemCatLogicalOrFilter;
            }
        }

        /// <summary>
        /// Общая функция для фильтра для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ElemExtraFilterFunc
        {
            get
            {
                if (_elemFilterFunc == null)
                {
                    _elemFilterFunc = (el) =>
                        el.Category != null
                        && !(el is ElementType)
                        // Молниезащита ЭОМ
                        && !(el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Полоса_")
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Пруток_")
                            || el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().StartsWith("Уголок_")
                            // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК
                            || (el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("ostec") && (el is FamilyInstance fi && fi.SuperComponent != null))
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("470_dkc_s5_accessories")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().ToLower().Contains("757_опора_неподвижная")
                            // Фильтрация семейств под которое НИКОГДА не должно быть отверстий
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("501_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("551_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("556_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("557_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("560_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("561_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("565_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("570_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("582_")
                            || el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString().StartsWith("592_"));
                }

                return _elemFilterFunc;
            }
        }


        /// <summary>
        /// Список BuiltInCategory для файлов ИОС, которые обрабатываются
        /// </summary>
        private static List<BuiltInCategory> BuiltInCategories
        {
            get =>
                new List<BuiltInCategory>()
                { 
                    // ОВВК (ЭОМСС - огнезащита)
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_DuctFitting,
                    BuiltInCategory.OST_DuctTerminal,
                    BuiltInCategory.OST_DuctAccessory,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_PipeFitting,
                    BuiltInCategory.OST_PipeAccessory,
                    BuiltInCategory.OST_MechanicalEquipment,
                    // ЭОМСС
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_CableTrayFitting,
                };
        }
    }
}
