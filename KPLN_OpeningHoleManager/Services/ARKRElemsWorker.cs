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
        /// <summary>
        /// Список BuiltInCategory ОТВЕРСТИЙ для файлов АР/КР, которые обрабатываются
        /// </summary>
        private static readonly BuiltInCategory[] _openingsBICColl = new BuiltInCategory[]
        {
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_Windows
        };

        /// <summary>
        /// Список BuiltInCategory ХОСТОВ для файлов АР/КР, которые обрабатываются
        /// </summary>
        private static readonly BuiltInCategory[] _hostBICColl = new BuiltInCategory[]
        {
            BuiltInCategory.OST_Walls,
        };

        private static LogicalOrFilter _openingElemCatLogicalOrFilter;
        private static LogicalOrFilter _hostElemCatLogicalOrFilter;
        private static Func<Element, bool> _openingFilterFunc;
        private static Func<Element, bool> _floorElemExtraFilterFunc;
        private static Func<Element, bool> _hostArElemFilterFunc;
        private static Func<Element, bool> _hostArkrElemFilterFunc;

        /// <summary>
        /// Список имен ХОСТОВ для файлов АР/КР, которые обрабатываются по правилу "начинается с"
        /// ВАЖНО: порядок в списке влияет на приоритет при выборе основ у соед. стен в методе "ClearCollectionByJoinedHosts"
        /// </summary>
        public static string[] ARKRHostNames_StartWith
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
        /// Список имен ХОСТОВ для файлов АР, которые обрабатываются по правилу "начинается с"
        /// </summary>
        public static string[] ARHostNames_StartWith
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
        /// Список имен ОТВЕРСТИЙ для файлов АР/КР, которые обрабатываются по правилу "начинается с"
        /// </summary>
        public static string[] OpeningNames_StartWith
        {
            get =>
                new string[]
                {
                    // Проекты КПЛН
                    "199_Отвер",
                    // Проекты СМЛТ
                    "ASML_АР_Отвер",
                };
        }

        /// <summary>
        /// Список имен ПЕРЕКРЫТИЙ для файлов АР/КР, которые обрабатываются по правилу "начинается с"
        /// </summary>
        public static string[] FloorNames_StartWith
        {
            get =>
                new string[]
                {
                    // Проекты КПЛН
                    "00_",
                    // Проекты СМЛТ
                    "КЖ_",
                };
        }

        /// <summary>
        /// Фильтр для ключевой фильтрации ХОСТОВ по категориям
        /// </summary>
        internal static LogicalOrFilter HostElemCatLogicalOrFilter
        {
            get
            {
                if (_hostElemCatLogicalOrFilter == null)
                {
                    List<ElementFilter> catFilters = new List<ElementFilter>();
                    catFilters.AddRange(_hostBICColl.Select(bic => new ElementCategoryFilter(bic)));

                    _hostElemCatLogicalOrFilter = new LogicalOrFilter(catFilters);
                }

                return _hostElemCatLogicalOrFilter;
            }
        }

        /// <summary>
        /// Фильтр для ключевой фильтрации ОТВЕРСТИЙ по категориям
        /// </summary>
        internal static LogicalOrFilter OpeningElemCatLogicalOrFilter
        {
            get
            {
                if (_openingElemCatLogicalOrFilter == null)
                {
                    List<ElementFilter> catFilters = new List<ElementFilter>();
                    catFilters.AddRange(_openingsBICColl.Select(bic => new ElementCategoryFilter(bic)));

                    _openingElemCatLogicalOrFilter = new LogicalOrFilter(catFilters);
                }

                return _openingElemCatLogicalOrFilter;
            }
        }

        /// <summary>
        /// Общая функция для фильтра ОТВЕРСТИЙ АР и КЖ для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> OpeningElemExtraFilterFunc
        {
            get
            {
                if (_openingFilterFunc == null)
                {
                    _openingFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        && OpeningNames_StartWith.Any(prefix => (bool)el.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)?.AsValueString()?.StartsWith(prefix));
                }

                return _openingFilterFunc;
            }
        }

        /// <summary>
        /// Общая функция для фильтра ОТВЕРСТИЙ АР и КЖ для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> FloorElemExtraFilterFunc
        {
            get
            {
                if (_floorElemExtraFilterFunc == null)
                {
                    _floorElemExtraFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        && FloorNames_StartWith.Any(prefix => (bool)el.Name?.StartsWith(prefix));
                }

                return _floorElemExtraFilterFunc;
            }
        }

        /// <summary>
        /// Общая функция для фильтра ХОСТОВ АР и КЖ для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ARKRHostElemExtraFilterFunc
        {
            get
            {
                if (_hostArkrElemFilterFunc == null)
                {
                    _hostArkrElemFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        // Анализируем стены АР и КР
                        && ARKRHostNames_StartWith.Any(prefix => (bool)el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix));
                }

                return _hostArkrElemFilterFunc;
            }
        }

        /// <summary>
        /// Общая функция для фильтра ХОСТОВ АР для ДОПОЛНИЕТЛЬНОГО просеивания элементов модели (новых и отредактированных)
        /// </summary>
        internal static Func<Element, bool> ARHostElemExtraFilterFunc
        {
            get
            {
                if (_hostArElemFilterFunc == null)
                {
                    _hostArElemFilterFunc = (el) =>
                        el.IsValidObject
                        && el.Category != null
                        && !(el is ElementType)
                        // Анализируем стены АР
                        && ARHostNames_StartWith.Any(prefix => (bool)el.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)?.AsValueString()?.StartsWith(prefix));
                }

                return _hostArElemFilterFunc;
            }
        }
    }
}
