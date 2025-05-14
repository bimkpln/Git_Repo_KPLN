using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Сервис по обработке эл-в АР и КР
    /// </summary>
    internal static class ARKRElemsWorker
    {
        private static LogicalOrFilter _elemCatLogicalOrFilter;
        private static Func<Element, bool> _elemFilterFunc;

        /// <summary>
        /// Список BuiltInCategory для файлов АР, которые обрабатываются
        /// </summary>
        public static BuiltInCategory[] BuiltInCategories
        {
            get =>
                new BuiltInCategory[]
                {
                    BuiltInCategory.OST_Walls,
                };
        }

        /// <summary>
        /// Список имен для файлов АР, которые обрабатываются по правилу "начинается с"
        /// </summary>
        public static string[] ARNames_StartWith
        {
            get =>
                new string[]
                {
                    "01_",
                    "02_",
                };
        }

        /// <summary>
        /// Список имен для файлов КЖ, которые обрабатываются по правилу "начинается с"
        /// </summary>
        public static string[] KRNames_StartWith
        {
            get =>
                new string[]
                {
                    "00_",
                };
        }

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
                        // Анализируем стены, кроме отделки
                        && ARNames_StartWith.Any(prefix => el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix) == true);
                }

                return _elemFilterFunc;
            }
        }
    }
}
