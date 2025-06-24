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
        private static Func<Element, bool> _arElemFilterFunc;
        private static Func<Element, bool> _arkrElemFilterFunc;

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
        /// Список имен для файлов АР/КР, которые обрабатываются по правилу "начинается с"
        /// ВАЖНО: порядок в списке влияет на приоритет при выборе основ у соед. стен в методе "ClearCollectionByJoinedHosts"
        /// </summary>
        public static string[] ARKRNames_StartWith
        {
            get =>
                new string[]
                {
                    // Проекты КПЛН
                    "00_",
                    "01_",
                    // Проекты СМЛТ
                    "КЖ_",
                    "ВС_",
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
                    // Утеплитель и фасады - можно было бы добавить, но они часто соед. с основной стеной АР, и это могут делать с ошибками в геом., что приводит к сложностям анализа
                    
                    // Проекты КПЛН
                    "01_",
                    // Проекты СМЛТ
                    "ВС_",
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
        /// Общая функция для фильтра АР и КЖ для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ARKRElemExtraFilterFunc
        {
            get
            {
                if (_arkrElemFilterFunc == null)
                {
                    _arkrElemFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        // Анализируем стены АР и КР
                        && ARKRNames_StartWith.Any(prefix => el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix) == true);
                }

                return _arkrElemFilterFunc;
            }
        }

        /// <summary>
        /// Общая функция для фильтра АР и КЖ для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ARElemExtraFilterFunc
        {
            get
            {
                if (_arElemFilterFunc == null)
                {
                    _arElemFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        // Анализируем стены АР
                        && ARNames_StartWith.Any(prefix => el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix) == true);
                }

                return _arElemFilterFunc;
            }
        }
    }
}
