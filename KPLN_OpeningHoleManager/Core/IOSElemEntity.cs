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

        internal IOSElemEntity(Document linkDoc, Element elem, Solid aRIOS_IntesectionSolid)
        {
            IOS_LinkDocument = linkDoc;
            IOS_Element = elem;
            ARKRIOS_IntesectionSolid = aRIOS_IntesectionSolid;
        }

        /// <summary>
        /// Ссылка на документ линка
        /// </summary>
        internal Document IOS_LinkDocument { get; private set; }

        /// <summary>
        /// Ссылка на элемент модели
        /// </summary>
        internal Element IOS_Element { get; private set; }

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
                    {
                        if (el.Category == null || el is ElementType)
                            return false;

                        string elem_type_param = el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.ToLower() ?? "";
                        string elem_family_param = el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString()?.ToLower() ?? "";
                        
                        return !(
                            // Молниезащита ЭОМ
                            elem_type_param.StartsWith("полоса_")
                            || elem_type_param.StartsWith("пруток_")
                            || elem_type_param.StartsWith("уголок_")
                            || elem_type_param.StartsWith("asml_эг_пруток-катанка")
                            || elem_type_param.StartsWith("asml_эг_полоса")
                            // Фильтрация семейств без геометрии от Ostec, крышка лотка DKC, неподвижную опору ОВВК
                            || (elem_family_param.Contains("ostec") && (el is FamilyInstance fi && fi.SuperComponent != null))
                            || elem_family_param.Contains("470_dkc_s5_accessories")
                            || elem_family_param.Contains("470_dkc_fireproof_out")
                            || elem_family_param.Contains("dkc_ceiling")
                            || elem_family_param.Contains("757_опора_неподвижная")
                            // Фильтрация семейств под которое НИКОГДА не должно быть отверстий
                            || elem_family_param.StartsWith("501_")
                            || elem_family_param.StartsWith("551_")
                            || elem_family_param.StartsWith("552_")
                            || elem_family_param.StartsWith("556_")
                            || elem_family_param.StartsWith("557_")
                            || elem_family_param.StartsWith("560_")
                            || elem_family_param.StartsWith("561_")
                            || elem_family_param.StartsWith("565_")
                            || elem_family_param.StartsWith("570_")
                            || elem_family_param.StartsWith("582_")
                            || elem_family_param.StartsWith("592_")
                            // Фильтрация типов семейств для которых опытным путём определено, что солид у них не взять (очень сложные семейства)
                            || elem_type_param.Contains("узел учета квартиры для гвс")
                            || elem_type_param.Contains("узел учета офиса для гвс")
                            || elem_type_param.Contains("узел учета квартиры для хвс")
                            || elem_type_param.Contains("узел учета офиса для хвс")
                            );
                    };
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
