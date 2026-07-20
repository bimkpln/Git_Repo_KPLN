using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.Common;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
#if R2023 || R2024
using static Autodesk.Revit.DB.SpecTypeId;
#endif
namespace KPLN_CoordiantorAI.ExternalModel
{
    internal class Commands
    {
        //public static object SpecTypeId { get; private set; }

        //private static Document _currentLinkedDoc = null;
        //private static Document _mainDoc = null;

        #region 1_get_active_view_in_revit

        public class ViewInfo
        {
            public long ViewId { get; set; }
            public string ViewName { get; set; }
            public string ViewType { get; set; }
            public bool IsSheet { get; set; }
        }

        public static ViewInfo GetActiveViewInfo(Autodesk.Revit.DB.Document doc)
        {
            View activeView = doc.ActiveView;

            return new ViewInfo
            {
                ViewId = IDHelper.ElIdInt(activeView.Id),
                ViewName = activeView.Name,
                ViewType = activeView.ViewType.ToString(),
                IsSheet = activeView is ViewSheet
            };
        }
        #endregion

        #region 2_get_all_elements_shown_in_view

        public class ViewElementsResult
        {
            public List<long> element_ids { get; set; } = new List<long>();
            public string view_name { get; set; }
            public string view_type { get; set; }
            public int count { get; set; }
        }

        public static ViewElementsResult GetAllElementsShownInView(Document doc, int viewOrSheetId)
        {
            ElementId viewId = IDHelper.ToElementId(viewOrSheetId);
            View view = doc.GetElement(viewId) as View;

            if (view == null)
            {
                return new ViewElementsResult
                {
                    element_ids = new List<long>(),
                    view_name = "Не найден",
                    view_type = "ERROR",
                    count = 0
                };
            }

            // Получаем ВСЕ элементы, видимые в этом виде
            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType();

            var elementIds = new List<long>();
            foreach (Element elem in collector)
            {
                elementIds.Add(IDHelper.ElIdInt(elem.Id));
            }

            return new ViewElementsResult
            {
                element_ids = elementIds,
                view_name = view.Name,
                view_type = view.ViewType.ToString(),
                count = elementIds.Count
            };
        }
        #endregion

        #region 4_get_category_by_keyword 

        public class CategoryInfo
        {
            public long Id { get; set; }
            public string Name { get; set; }
            public BuiltInCategory BuiltInCategory { get; set; }
        }

        public static List<CategoryInfo> GetCategoryByKeyword(Document doc, string keyword)
        {
            var matchingCategories = new List<CategoryInfo>();

            // Получаем все категории документа
            Categories categories = doc.Settings.Categories;

            foreach (Category category in categories)
            {
                if (category == null) continue;

                string categoryName = category.Name ?? "";

                // Ищем по ключевому слову (регистронезависимо)
                if (categoryName.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    matchingCategories.Add(new CategoryInfo
                    {
                        Id = IDHelper.ElIdInt(category.Id),
                        Name = categoryName,
                        BuiltInCategory = (BuiltInCategory)IDHelper.ElIdInt(category.Id)
                    });
                }
            }

            return matchingCategories;
        }
        #endregion

        #region 3_get_elements_by_category


        public static List<long> GetElementsByCategory(Document doc, int categoryId)
        {
            var elementIds = new List<long>();

            // Фильтр по категории
            ElementCategoryFilter categoryFilter = new ElementCategoryFilter((BuiltInCategory)categoryId);
            IList<Element> elements = new FilteredElementCollector(doc)
                .WherePasses(categoryFilter)
                .WhereElementIsNotElementType()  // Только экземпляры, не типы
                .ToElements();

            foreach (Element elem in elements)
            {
                elementIds.Add(IDHelper.ElIdInt(elem.Id));
            }

            return elementIds;
        }

        #endregion

        #region 5_get_model_categories


        public static List<CategoryInfo> GetModelCategories(Document doc)
        {
            var result = new List<CategoryInfo>();

            Categories categories = doc.Settings.Categories;
            foreach (Category cat in categories)
            {
                if (cat == null) continue;
                if (string.IsNullOrWhiteSpace(cat.Name)) continue;

                result.Add(new CategoryInfo
                {
                    Id = IDHelper.ElIdInt(cat.Id),
                    Name = cat.Name
                });
            }

            // можно отсортировать по имени для удобства
            return result
                .OrderBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        #endregion

        #region 6_get_categories_from_elementids   


        public static Dictionary<int, CategoryInfo> GetCategoriesFromElementIds(Document doc, IEnumerable<int> elementIds)
        {
            var result = new Dictionary<int, CategoryInfo>();

            if (elementIds == null)
                return result;

            foreach (int idInt in elementIds.Distinct())
            {
                var eid = IDHelper.ToElementId(idInt);
                Element elem = doc.GetElement(eid);
                if (elem == null)
                    continue;

                Category cat = elem.Category;
                if (cat == null)
                    continue;

                result[idInt] = new CategoryInfo
                {
                    Id = IDHelper.ElIdInt(cat.Id),
                    Name = cat.Name
                };
            }

            return result;
        }

        #endregion

        #region 7_get_element_types_for_elementids


        public class ElementTypeInfo
        {
            public int TypeId { get; set; }
            public string TypeName { get; set; }
        }
        public static object GetElementTypesForElementIds(Document doc, IEnumerable<int> elementIds)
        {
            var typeMap = new Dictionary<int, ElementTypeInfo>();
            int processedCount = 0;

            if (elementIds == null) return new { type_ids = typeMap, count = 0 };

            foreach (int idInt in elementIds.Distinct())
            {
                var eid = IDHelper.ToElementId(idInt);
                Element elem = doc.GetElement(eid);
                if (elem == null) continue;

                processedCount++;

                // Получаем тип элемента
                ElementId typeId = elem.GetTypeId();
                Element typeElem = doc.GetElement(typeId);

                if (typeElem != null)
                {
                    string typeName = typeElem.Name ?? "Без имени";
                    typeMap[idInt] = new ElementTypeInfo
                    {
                        TypeId = IDHelper.ElIdInt(typeId),
                        TypeName = typeName
                    };
                }
            }

            return new
            {
                type_ids = typeMap,
                count = processedCount
            };
        }


        #endregion

        #region 8_get_all_elementids_for_specific_type_ids


        public static object GetAllElementIdsForSpecificTypeIds(Document doc, IEnumerable<int> typeIds)
        {
            var elementIdsPerType = new Dictionary<int, List<int>>();
            int processedTypes = 0;

            if (typeIds == null)
                return new { element_ids_per_type = elementIdsPerType, processed_types = 0 };

            foreach (int typeIdInt in typeIds.Distinct())
            {
                processedTypes++;
                var typeId = IDHelper.ToElementId(typeIdInt);
                var elemIds = new List<int>();

                // Фильтр по TypeId
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();  // Только экземпляры

                foreach (Element elem in collector)
                {
                    if (IDHelper.ElIdInt(elem.GetTypeId()) == typeIdInt)
                        elemIds.Add(IDHelper.ElIdInt(elem.Id));
                }

                elementIdsPerType[typeIdInt] = elemIds;
            }

            return new
            {
                element_ids_per_type = elementIdsPerType,
                processed_types = processedTypes
            };
        }


        #endregion

        #region 9_get_all_used_families_in_model   

        public class FamilyInfo
        {
            public int FamilyId { get; set; }
            public string FamilyName { get; set; }
            public bool IsLoadedFamily { get; set; }      // true = загружаемое семейство, false = системное
            public bool IsPlacedInModel { get; set; }     // true = размещено в модели, false = не размещено
            public int InstanceCount { get; set; }        // Количество экземпляров
        }

        public static object GetAllUsedFamiliesInModel(Document doc)
        {
            var families = new List<FamilyInfo>();
            var existingLoadedIds = new HashSet<int>();        // ID загружаемых семейств
            var existingSystemNames = new HashSet<string>();  // Имена системных семейств

            // ========== 1. Собираем ЗАГРУЖАЕМЫЕ семейства ==========
            var loadedFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            // Собирает ВСЕ экземпляры элементов в документе (кроме типов)
            var allElements = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements();


            // ========== 3. Подсчёт экземпляров для загружаемых семейств ==========
            var loadedInstanceCountMap = new Dictionary<int, int>();
            foreach (var elem in allElements)
            {
                FamilyInstance fi = elem as FamilyInstance;
                if (fi != null && fi.Symbol != null && fi.Symbol.Family != null)
                {
                    int familyId = IDHelper.ElIdInt(fi.Symbol.Family.Id);
                    if (loadedInstanceCountMap.ContainsKey(familyId))
                        loadedInstanceCountMap[familyId]++;
                    else
                        loadedInstanceCountMap[familyId] = 1;
                }
            }

            // ========== 4. Добавляем загружаемые семейства ==========
            foreach (Family family in loadedFamilies)
            {
                int familyId = IDHelper.ElIdInt(family.Id);

                if (existingLoadedIds.Contains(familyId))
                    continue;

                existingLoadedIds.Add(familyId);

                bool isPlacedInModel = loadedInstanceCountMap.ContainsKey(familyId) && loadedInstanceCountMap[familyId] > 0;
                int instanceCount = isPlacedInModel ? loadedInstanceCountMap[familyId] : 0;

                families.Add(new FamilyInfo
                {
                    FamilyId = familyId,
                    FamilyName = family.Name ?? "Unnamed",
                    IsLoadedFamily = true,
                    IsPlacedInModel = isPlacedInModel,
                    InstanceCount = instanceCount
                });
            }


            // ========== 5. Подсчёт экземпляров для системных семейств ==========
            var systemInstanceCountMap = new Dictionary<string, int>();
            foreach (var elem in allElements)
            {
                // Пытаемся привести к FamilyInstance
                FamilyInstance fi = elem as FamilyInstance;
                if (fi != null)
                {
                    // Это ЗАГРУЖАЕМОЕ семейство
                    if (fi.Symbol != null && fi.Symbol.Family != null)
                    {
                        continue;
                    }
                }
                else
                {
                    // Получаем имя системного семейства
                    ElementType elemType = doc.GetElement(elem.GetTypeId()) as ElementType;

                    string familyName = elemType?.FamilyName;
                    if (string.IsNullOrEmpty(familyName))
                    {
                        familyName = elem.Category?.Name ?? "Unknown";
                    }

                    if (!string.IsNullOrEmpty(familyName))
                    {
                        if (systemInstanceCountMap.ContainsKey(familyName))
                            systemInstanceCountMap[familyName]++;
                        else
                            systemInstanceCountMap[familyName] = 1;
                    }
                }
            }

            // ========== 6. Собираем системные семейства через ElementType ==========
            var allElementTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .ToList();

            foreach (ElementType elemType in allElementTypes)
            {
                // Проверяем, является ли тип загружаемым семейством
                FamilySymbol symbol = elemType as FamilySymbol;
                bool isLoadedFamilyType = false;

                if (symbol != null)
                {
                    try
                    {
                        isLoadedFamilyType = (symbol.Family != null && IDHelper.ElIdInt(symbol.Family.Id) != -1); //значит с-во загружаемое
                    }
                    catch
                    {
                        isLoadedFamilyType = false; //значит с-во системное
                    }
                }

                if (isLoadedFamilyType)
                    continue; // Пропускаем загружаемые семейства (уже обработаны)

                // Получаем имя системного семейства
                string familyName = elemType?.FamilyName;
                if (string.IsNullOrEmpty(familyName))
                {
                    familyName = elemType.Category?.Name ?? "Unknown";
                }

                // Проверяем, не добавили ли уже такое системное семейство
                if (existingSystemNames.Contains(familyName))
                    continue;

                // Фильтруем служебные типы
                if (!IsRelevantSystemFamily(elemType))
                    continue;

                existingSystemNames.Add(familyName);

                bool isPlacedInModel = systemInstanceCountMap.ContainsKey(familyName) && systemInstanceCountMap[familyName] > 0;
                int instanceCount = isPlacedInModel ? systemInstanceCountMap[familyName] : 0;

                families.Add(new FamilyInfo
                {
                    FamilyId = -1,//-Math.Abs(familyName.GetHashCode()), // Отрицательный ID на основе хэша имени
                    FamilyName = familyName,
                    IsLoadedFamily = false,
                    IsPlacedInModel = isPlacedInModel,
                    InstanceCount = instanceCount
                });
            }

            // ========== 7. Сортируем и возвращаем результат ==========
            return new
            {
                families = families.OrderBy(f => f.FamilyName).ToList(),
                count = families.Count,
                stats = new
                {
                    loaded_families = families.Count(f => f.IsLoadedFamily),
                    system_families = families.Count(f => !f.IsLoadedFamily),
                    placed_families = families.Count(f => f.IsPlacedInModel),
                    unplaced_families = families.Count(f => !f.IsPlacedInModel)
                }
            };
        }

        /// <summary>
        /// Проверяет, является ли тип элемента релевантным системным семейством
        /// (исключает типы размеров, материалов и другие служебные типы)
        /// </summary>
        private static bool IsRelevantSystemFamily(ElementType elemType)
        {
            // Исключаем типы, которые не являются "семействами" в понимании пользователя
            if (elemType is DimensionType) return false;
            if (elemType is Material) return false;
            if (elemType is GraphicsStyle) return false;
            if (elemType is FillPatternElement) return false;

            return true;
        }


        #endregion

        #region 10_get_all_used_families_of_category
        public static object GetAllUsedFamiliesOfCategory(Document doc, int categoryId)
        {
            var families = new List<FamilyInfo>();

            //все элементы в модели
            var allElements = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .ToElements();

            //все загружаемые с-ва
            var loadedFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            //все типоразмеры (нужно для системных с-в)
            var allElementTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>()
                .ToList();

            //Создание словаря для загржаемых семейств
            var loadedInstanceCount = new Dictionary<int, int>();
            foreach (var elem in allElements)
            {
                FamilyInstance fi = elem as FamilyInstance;
                if (fi != null && fi.Symbol != null && fi.Symbol.Family != null)
                {
                    int fId = IDHelper.ElIdInt(fi.Symbol.Family.Id);
                    if (loadedInstanceCount.ContainsKey(fId))
                        loadedInstanceCount[fId]++;
                    else
                        loadedInstanceCount[fId] = 1;
                }
            }

            //Создание словаря для системных семейств
            var systemInstanceCount = new Dictionary<string, int>();
            foreach (var elem in allElements)
            {
                string fname = (doc.GetElement(elem.GetTypeId()) as ElementType)?.FamilyName;
                if (string.IsNullOrEmpty(fname))
                    fname = elem.Category?.Name ?? "Unknown";

                if (!string.IsNullOrEmpty(fname))
                {
                    if (systemInstanceCount.ContainsKey(fname))
                        systemInstanceCount[fname]++;
                    else
                        systemInstanceCount[fname] = 1;
                }
            }



            //ОБРАБОТКА ЗАГРУЖАЕМЫХ СЕМЕЙСТВ
            foreach (var inst in loadedFamilies)
            {
                if (inst.FamilyCategory == null || IDHelper.ElIdInt(inst.FamilyCategory.Id) != categoryId)
                    continue;

                //Получаем ID
                int familyId = IDHelper.ElIdInt(inst.Id);

                if (families.Any(f => f.FamilyId == familyId))
                    continue;

                //Получаем имя
                string familyName = inst.Name;

                //получаем есть ли экземпляр в модели, если да то в каком количестве
                bool isPlacedInModel = loadedInstanceCount.ContainsKey(familyId) && loadedInstanceCount[familyId] > 0;
                int instanceCount = isPlacedInModel ? loadedInstanceCount[familyId] : 0;

                //Добавляем в коллекцию загружаемые с-ва
                families.Add(new FamilyInfo
                {
                    FamilyId = familyId,
                    FamilyName = familyName ?? "Unnamed",
                    IsLoadedFamily = true,
                    IsPlacedInModel = isPlacedInModel,
                    InstanceCount = instanceCount
                });
            }

            //ОБРАБОТКА СИСТЕМНЫХ СЕМЕЙСТВ
            foreach (var syselemType in allElementTypes)
            {
                if (syselemType.Category == null || IDHelper.ElIdInt(syselemType.Category.Id) != categoryId)
                    continue;

                // Проверяем, является ли типоразмер загружаемым семейством
                FamilySymbol symbol = syselemType as FamilySymbol;
                if (symbol != null && symbol.Family != null && IDHelper.ElIdInt(symbol.Family.Id) != -1)
                    continue;

                //Получаем ID
                int familyId = -1;//-Math.Abs(familyName.GetHashCode());

                //Получаем имя
                string familyName = syselemType.FamilyName;

                //получаем есть ли экземпляр в модели, если да то в каком количестве
                bool isPlacedInModel = systemInstanceCount.ContainsKey(familyName) && systemInstanceCount[familyName] > 0;
                int instanceCount = isPlacedInModel ? systemInstanceCount[familyName] : 0;

                families.Add(new FamilyInfo
                {
                    FamilyId = familyId,
                    FamilyName = familyName ?? "Unnamed",
                    IsLoadedFamily = false,
                    IsPlacedInModel = isPlacedInModel,
                    InstanceCount = instanceCount
                });
            }


            return new
            {
                families = families.OrderBy(f => f.FamilyName).ToList(),
                count = families.Count,
                stats = new
                {
                    loaded_families = families.Count(f => f.IsLoadedFamily),
                    system_families = families.Count(f => !f.IsLoadedFamily),
                    placed_families = families.Count(f => f.IsPlacedInModel),
                    unplaced_families = families.Count(f => !f.IsPlacedInModel)
                }
            };
        }

        // Чтобы Distinct работал по Id
        public class FamilyIdComparer : IEqualityComparer<Family>
        {
            public bool Equals(Family x, Family y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (x is null || y is null) return false;
                return IDHelper.ElIdInt(x.Id) == IDHelper.ElIdInt(y.Id);
            }

            public int GetHashCode(Family obj)
            {
                return IDHelper.ElIdInt(obj.Id);
            }
        }

        #endregion

        #region 11_get_all_used_types_of_a_family   


        public class FamilyTypeInfo
        {
            public int TypeId { get; set; }
            public string TypeName { get; set; }
        }

        public static object GetAllUsedTypesOfAFamily(Document doc, string familyName)
        {
            var types = new List<FamilyTypeInfo>();

            if (string.IsNullOrWhiteSpace(familyName))
                return new { types, count = 0 };

            // Берём ВСЕ типы в документе (включая системные: стены, перекрытия и т.д.)
            var allTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>();

            // Фильтруем по точному имени семейства (чувствительно к регистру)
            var familyTypes = allTypes
                .Where(t => t.FamilyName == familyName)   // <-- КЛЮЧЕВАЯ СТРОКА
                .ToList();

            foreach (var t in familyTypes)
            {
                types.Add(new FamilyTypeInfo
                {
                    TypeId = IDHelper.ElIdInt(t.Id),
                    TypeName = t.Name
                });
            }

            return new
            {
                types = types.OrderBy(x => x.TypeName).ToList(),
                count = types.Count
            };
        }


        #endregion

        #region 12_get_all_elements_of_specific_families 


        public class ElementsOfFamilyInfo
        {
            public string FamilyName { get; set; }
            public List<int> ElementIds { get; set; }
        }

        public static object GetAllElementsOfSpecificFamilies(Document doc, List<string> familyNames)
        {
            var elementsPerFamily = new Dictionary<string, List<int>>();

            if (familyNames == null || familyNames.Count == 0)
                return new { elements_per_family = elementsPerFamily };

            // Для быстрых точных совпадений по имени семейства
            var familyNameSet = new HashSet<string>(familyNames);

            // 1) Берём ВСЕ типы во всём проекте
            var allTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .Cast<ElementType>();

            // 2) Отбираем только типы нужных семейств по точному FamilyName
            var typesOfFamilies = allTypes
                .Where(t => !string.IsNullOrEmpty(t.FamilyName))
                .Where(t => familyNameSet.Contains(t.FamilyName))
                .ToList();

            // 3) Для каждого типа — находим все экземпляры этого типа
            foreach (var type in typesOfFamilies)
            {
                string familyName = type.FamilyName;
                if (!elementsPerFamily.TryGetValue(familyName, out var list))
                {
                    list = new List<int>();
                    elementsPerFamily[familyName] = list;
                }

                // Коллектор только по этому TypeId
                var instancesOfType = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .Where(e => IDHelper.ElIdInt(e.GetTypeId()) == IDHelper.ElIdInt(type.Id));

                foreach (var inst in instancesOfType)
                    list.Add(IDHelper.ElIdInt(inst.Id));
            }

            return new
            {
                elements_per_family = elementsPerFamily
            };
        }


        #endregion

        #region 13_get_parameters_from_elementid   


        public class ElementParameterInfo
        {
            public int ParameterId { get; set; }
            public string ParameterName { get; set; }
            public string Value { get; set; }
            public string StorageType { get; set; }
            public bool IsReadOnly { get; set; }
            public string parType { get; set; }
            public string Error { get; set; }
        }

        public static object GetParametersFromElementId(Document doc, int elementId, bool getIdValuesAsNames)
        {
            var result = new List<ElementParameterInfo>();
            var errors = new List<string>();

            var elem = doc.GetElement(IDHelper.ToElementId(elementId));
            if (elem == null)
                return new { parameters = result };

            //параметры экземпляра
            foreach (Parameter p in elem.Parameters)
            {
                if (p == null) continue;
                try
                {

                    string name = p.Definition?.Name ?? "<no name>";
                string storageType = p.StorageType.ToString();
                bool isReadOnly = p.IsReadOnly;

                string value = "";
                switch (p.StorageType)
                {
                    case StorageType.String:
                        value = p.AsString();
                        break;

                    case StorageType.Double:
                        // Оставляем в “сыром” виде, без UnitUtils, чтобы не навязывать единицы
                        value = p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                        break;

                    case StorageType.Integer:
                        value = p.AsInteger().ToString();
                        break;

                    case StorageType.ElementId:
                        var id = p.AsElementId();
                        if (id != ElementId.InvalidElementId)
                        {
                            if (getIdValuesAsNames)
                            {
                                var refElem = doc.GetElement(id);
                                value = refElem != null
                                    ? refElem.Name
                                    : IDHelper.ElIdInt(id).ToString();
                            }
                            else
                            {
                                value = IDHelper.ElIdInt(id).ToString();
                            }
                        }
                        else
                        {
                            value = "InvalidElementId";
                        }
                        break;

                    case StorageType.None:
                    default:
                        value = p.AsValueString(); // попытка взять красиво отформатированное значение
                        break;
                }

                // У Definition нет “Id” как у параметров семейства, поэтому используем Id самого параметра
                int paramId = p.Id != null ? IDHelper.ElIdInt(p.Id) : 0;

                result.Add(new ElementParameterInfo
                {
                    ParameterId = paramId,
                    ParameterName = name,
                    Value = value,
                    StorageType = storageType,
                    IsReadOnly = isReadOnly,
                    parType = "exemplar"
                });
                }
                catch (Exception ex)
                {
                    errors.Add("Failed to read exemplar parameter: " + ex.Message);
                }

            }

            var elemType = doc.GetElement(elem.GetTypeId());

            //параметры типоразмера
            if (elemType != null)
            {
                foreach (Parameter p in elemType.Parameters)
                {
                    if (p == null) continue;
                    try
                    {

                        string name = p.Definition?.Name ?? "<no name>";
                        string storageType = p.StorageType.ToString();
                    bool isReadOnly = p.IsReadOnly;

                    string value = "";
                    switch (p.StorageType)
                    {
                        case StorageType.String:
                            value = p.AsString();
                            break;

                        case StorageType.Double:
                            // Оставляем в “сыром” виде, без UnitUtils, чтобы не навязывать единицы
                            value = p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                            break;

                        case StorageType.Integer:
                            value = p.AsInteger().ToString();
                            break;

                        case StorageType.ElementId:
                            var id = p.AsElementId();
                            if (id != ElementId.InvalidElementId)
                            {
                                if (getIdValuesAsNames)
                                {
                                    var refElem = doc.GetElement(id);
                                    value = refElem != null
                                        ? refElem.Name
                                        : IDHelper.ElIdInt(id).ToString();
                                }
                                else
                                {
                                    value = IDHelper.ElIdInt(id).ToString();
                                }
                            }
                            else
                            {
                                value = "InvalidElementId";
                            }
                            break;

                        case StorageType.None:
                        default:
                            value = p.AsValueString(); // попытка взять красиво отформатированное значение
                            break;
                    }

                    // У Definition нет “Id” как у параметров семейства, поэтому используем Id самого параметра
                    int paramId = p.Id != null ? IDHelper.ElIdInt(p.Id) : 0;

                    result.Add(new ElementParameterInfo
                    {
                        ParameterId = paramId,
                        ParameterName = name,
                        Value = value,
                        StorageType = storageType,
                        IsReadOnly = isReadOnly,
                        parType = "Type"
                    });
                    }
                    catch (Exception ex)
                    {
                        errors.Add("Failed to read type parameter: " + ex.Message);
                    }
                }
            }
            else
            {
                errors.Add("Type element was not found for element type id.");
            }


            return new
            {
                parameters = result,
                errors = errors
            };
        }


        #endregion

        #region 14_get_parameter_value_for_element_ids


        public static object GetParameterValueForElementIds(Document doc, List<int> elementIds, int idParameter, bool getIdValuesAsNames)
        {
            var values = new Dictionary<int, string>();

            if (elementIds == null || elementIds.Count == 0)
                return new { parameter_values = values };

            foreach (int eidInt in elementIds.Distinct())
            {
                var elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null)
                {
                    values[eidInt] = null;
                    continue;
                }

                // Ищем параметр по Id
                Parameter p = null;
                foreach (Parameter param in elem.Parameters)
                {
                    if (param?.Id != null && IDHelper.ElIdInt(param.Id) == idParameter)
                    {
                        p = param;
                        break;
                    }
                }

                if (p == null)
                {
                    values[eidInt] = null;
                    continue;
                }

                string val = "";
                switch (p.StorageType)
                {
                    case StorageType.String:
                        val = p.AsString();
                        break;

                    case StorageType.Double:
                        val = p.AsDouble()
                            .ToString(System.Globalization.CultureInfo.InvariantCulture);
                        break;

                    case StorageType.Integer:
                        val = p.AsInteger().ToString();
                        break;

                    case StorageType.ElementId:
                        var id = p.AsElementId();
                        if (id != ElementId.InvalidElementId)
                        {
                            if (getIdValuesAsNames)
                            {
                                var refElem = doc.GetElement(id);
                                val = refElem != null
                                    ? refElem.Name
                                    : IDHelper.ElIdInt(id).ToString();
                            }
                            else
                            {
                                val = IDHelper.ElIdInt(id).ToString();
                            }
                        }
                        else
                        {
                            val = "InvalidElementId";
                        }
                        break;

                    case StorageType.None:
                    default:
                        val = p.AsValueString();
                        break;
                }

                values[eidInt] = val;
            }

            return new
            {
                parameter_values = values
            };
        }


        #endregion

        #region 15_get_all_additional_properties_from_elementid  


        public class AdditionalPropertyInfo
        {
            public string Name { get; set; }
            public string Value { get; set; }
        }

        public static object GetAllAdditionalPropertiesFromElementId(Document doc, int elementId)
        {
            var result = new List<AdditionalPropertyInfo>();

            var elem = doc.GetElement(IDHelper.ToElementId(elementId));
            if (elem == null)
                return new { additional_properties = result };

            void AddProp(string name, object value)
            {
                if (value == null) return;
                result.Add(new AdditionalPropertyInfo
                {
                    Name = name,
                    Value = value.ToString()
                });
            }

            // Общие свойства Element
            AddProp("Id", IDHelper.ElIdInt(elem.Id));
            AddProp("UniqueId", elem.UniqueId);
            AddProp("Name", elem.Name);
            AddProp("Category", elem.Category?.Name);
            AddProp("LevelId", elem.LevelId != null ? IDHelper.ElIdInt(elem.LevelId) : (int?)null);
            AddProp("OwnerViewId", elem.OwnerViewId != null ? IDHelper.ElIdInt(elem.OwnerViewId) : (int?)null);

            var loc = elem.Location;
            if (loc is LocationPoint lp)
            {
                AddProp("LocationPoint_X", lp.Point.X);
                AddProp("LocationPoint_Y", lp.Point.Y);
                AddProp("LocationPoint_Z", lp.Point.Z);
            }
            else if (loc is LocationCurve lc)
            {
                AddProp("LocationCurve_Length", lc.Curve.Length);
            }

            var bbox = elem.get_BoundingBox(null);
            if (bbox != null)
            {
                AddProp("BoundingBox_Min", bbox.Min);
                AddProp("BoundingBox_Max", bbox.Max);
            }

            // Специализированные свойства по типу элемента
            if (elem is Wall wall)
            {
                AddProp("Wall_Width", wall.Width);
                AddProp("Wall_Kind", wall.WallType?.Kind);
            }
            else if (elem is Floor floor)
            {
                AddProp("FloorType_Name", floor.FloorType?.Name);
            }
            else if (elem is FamilyInstance fi)
            {
                AddProp("FamilyName", fi.Symbol?.FamilyName);
                AddProp("SymbolName", fi.Symbol?.Name);
                AddProp("HostId", fi.Host != null ? IDHelper.ElIdInt(fi.Host.Id) : (int?)null);
                AddProp("RoomId", fi.Room != null ? IDHelper.ElIdInt(fi.Room.Id) : (int?)null);
                AddProp("SpaceId", fi.Space != null ? IDHelper.ElIdInt(fi.Space.Id) : (int?)null);
            }
            else if (elem is Room room)
            {
                AddProp("RoomNumber", room.Number);
                AddProp("RoomName", room.Name);
                AddProp("RoomArea", room.Area);
                AddProp("RoomLevelId", IDHelper.ElIdInt(room.LevelId));
            }

            return new
            {
                additional_properties = result
            };
        }


        #endregion

        #region 16_get_additional_property_for_all_elementids   


        public static object GetAdditionalPropertyForAllElementIds(Document doc, List<int> elementIds, string propertyName)
        {
            var propertyValues = new Dictionary<int, string>();
            var invalidElementIds = new List<int>();

            if (elementIds == null || elementIds.Count == 0 || string.IsNullOrWhiteSpace(propertyName))
                return new { property_values = propertyValues, invalid_element_ids = invalidElementIds };

            foreach (int eidInt in elementIds.Distinct())
            {
                var elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null)
                {
                    invalidElementIds.Add(eidInt);
                    continue;
                }

                string value = null;
                bool found = false;

                // Общие свойства Element
                switch (propertyName)
                {
                    case "Id":
                        value = IDHelper.ElIdInt(elem.Id).ToString();
                        found = true;
                        break;

                    case "UniqueId":
                        value = elem.UniqueId;
                        found = true;
                        break;

                    case "Name":
                        value = elem.Name;
                        found = true;
                        break;

                    case "Category":
                        value = elem.Category?.Name;
                        found = true;
                        break;

                    case "LevelId":
                        value = elem.LevelId != null ? IDHelper.ElIdInt(elem.LevelId).ToString() : null;
                        found = true;
                        break;

                    case "OwnerViewId":
                        value = elem.OwnerViewId != null ? IDHelper.ElIdInt(elem.OwnerViewId).ToString() : null;
                        found = true;
                        break;
                }

                // Location / BoundingBox и спец‑типы — только если ещё не нашли
                if (!found)
                {
                    var loc = elem.Location;
                    if (loc is LocationPoint lp)
                    {
                        switch (propertyName)
                        {
                            case "LocationPoint_X":
                                value = lp.Point.X.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                found = true;
                                break;
                            case "LocationPoint_Y":
                                value = lp.Point.Y.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                found = true;
                                break;
                            case "LocationPoint_Z":
                                value = lp.Point.Z.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                found = true;
                                break;
                        }
                    }
                    else if (loc is LocationCurve lc)
                    {
                        if (propertyName == "LocationCurve_Length")
                        {
                            value = lc.Curve.Length.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            found = true;
                        }
                    }
                }

                if (!found)
                {
                    var bbox = elem.get_BoundingBox(null);
                    if (bbox != null)
                    {
                        switch (propertyName)
                        {
                            case "BoundingBox_Min":
                                value = bbox.Min.ToString();
                                found = true;
                                break;
                            case "BoundingBox_Max":
                                value = bbox.Max.ToString();
                                found = true;
                                break;
                        }
                    }
                }

                // Специализированные классы
                if (!found && elem is Wall wall)
                {
                    switch (propertyName)
                    {
                        case "Wall_Width":
                            value = wall.Width.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            found = true;
                            break;
                        case "Wall_Kind":
                            value = wall.WallType?.Kind.ToString();
                            found = true;
                            break;
                    }
                }
                else if (!found && elem is Floor floor)
                {
                    if (propertyName == "FloorType_Name")
                    {
                        value = floor.FloorType?.Name;
                        found = true;
                    }
                }
                else if (!found && elem is FamilyInstance fi)
                {
                    switch (propertyName)
                    {
                        case "FamilyName":
                            value = fi.Symbol?.FamilyName;
                            found = true;
                            break;
                        case "SymbolName":
                            value = fi.Symbol?.Name;
                            found = true;
                            break;
                        case "HostId":
                            value = fi.Host != null ? IDHelper.ElIdInt(fi.Host.Id).ToString() : null;
                            found = true;
                            break;
                        case "RoomId":
                            value = fi.Room != null ? IDHelper.ElIdInt(fi.Room.Id).ToString() : null;
                            found = true;
                            break;
                        case "SpaceId":
                            value = fi.Space != null ? IDHelper.ElIdInt(fi.Space.Id).ToString() : null;
                            found = true;
                            break;
                    }
                }
                else if (!found && elem is Room room)
                {
                    switch (propertyName)
                    {
                        case "RoomNumber":
                            value = room.Number;
                            found = true;
                            break;
                        case "RoomName":
                            value = room.Name;
                            found = true;
                            break;
                        case "RoomArea":
                            value = room.Area.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            found = true;
                            break;
                        case "RoomLevelId":
                            value = IDHelper.ElIdInt(room.LevelId).ToString();
                            found = true;
                            break;
                    }
                }

                if (found)
                    propertyValues[eidInt] = value;
                else
                    invalidElementIds.Add(eidInt);
            }

            return new
            {
                property_values = propertyValues,
                invalid_element_ids = invalidElementIds
            };
        }


        #endregion

        #region 16_1_get_revitlookup_like_properties

        public static object GetRevitLookupLikeProperties(
            Document doc,
            int elementId,
            int maxValueLength = 1000)
        {
            try
            {
                Element elem = doc.GetElement(IDHelper.ToElementId(elementId));
                if (elem == null)
                {
                    return new
                    {
                        success = false,
                        error = $"Элемент с ID {elementId} не найден.",
                        element_id = elementId
                    };
                }

                int safeMaxValueLength = Math.Max(100, Math.Min(maxValueLength, 10000));
                var apiProperties = new List<object>();
                var specialProperties = new List<object>();


                PropertyInfo[] properties = elem.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
                    foreach (PropertyInfo property in properties.OrderBy(p => p.Name))
                    {
                        if (!property.CanRead || property.GetIndexParameters().Length > 0)
                            continue;

                        try
                        {
                            object value = property.GetValue(elem, null);
                            apiProperties.Add(new
                            {
                                name = property.Name,
                                type = property.PropertyType.FullName,
                                value = TrimForLookup(FormatLookupValue(value), safeMaxValueLength)
                            });
                        }
                        catch (Exception ex)
                        {
                            apiProperties.Add(new
                            {
                                name = property.Name,
                                type = property.PropertyType.FullName,
                                error = ex.Message
                            });
                        }
                    }

                AddSpecialLookupProperties(doc, elem, specialProperties, safeMaxValueLength);
                return new
                {
                    success = true,
                    element_id = elementId,
                    unique_id = elem.UniqueId,
                    class_name = elem.GetType().FullName,
                    category = elem.Category != null ? elem.Category.Name : null,
                    name = elem.Name,
                    api_properties_count = apiProperties.Count,
                    api_properties = apiProperties,
                    special_properties = specialProperties
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    success = false,
                    error = $"Ошибка при получении свойств элемента: {ex.Message}",
                    element_id = elementId
                };
            }
        }

        private static string FormatParameterValue(Document doc, Parameter parameter)
        {
            if (parameter == null)
                return null;

            switch (parameter.StorageType)
            {
                case StorageType.String:
                    return parameter.AsString();

                case StorageType.Double:
                    return parameter.AsDouble().ToString(CultureInfo.InvariantCulture);

                case StorageType.Integer:
                    return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);

                case StorageType.ElementId:
                    ElementId id = parameter.AsElementId();
                    if (id == ElementId.InvalidElementId)
                        return "InvalidElementId";

                    Element refElement = doc.GetElement(id);
                    string idText = IDHelper.ElIdInt(id).ToString(CultureInfo.InvariantCulture);
                    return refElement == null ? idText : $"{idText} ({refElement.Name})";

                case StorageType.None:
                default:
                    return parameter.AsValueString();
            }
        }

        private static string FormatLookupValue(object value)
        {
            if (value == null)
                return null;

            if (value is ElementId elementId)
                return IDHelper.ElIdInt(elementId).ToString(CultureInfo.InvariantCulture);

            if (value is Element element)
                return $"{element.GetType().Name}: {element.Name} [{IDHelper.ElIdInt(element.Id)}]";

            if (value is Category category)
                return $"{category.Name} [{IDHelper.ElIdInt(category.Id)}]";

            if (value is XYZ xyz)
                return $"X={xyz.X.ToString(CultureInfo.InvariantCulture)}, Y={xyz.Y.ToString(CultureInfo.InvariantCulture)}, Z={xyz.Z.ToString(CultureInfo.InvariantCulture)}";

            if (value is BoundingBoxXYZ bbox)
                return $"Min=({FormatLookupValue(bbox.Min)}), Max=({FormatLookupValue(bbox.Max)})";

            if (value is Transform transform)
                return $"Origin=({FormatLookupValue(transform.Origin)}), BasisX=({FormatLookupValue(transform.BasisX)}), BasisY=({FormatLookupValue(transform.BasisY)}), BasisZ=({FormatLookupValue(transform.BasisZ)})";

            if (value is System.Collections.IEnumerable enumerable && !(value is string))
            {
                var items = new List<string>();
                int count = 0;
                foreach (object item in enumerable)
                {
                    if (count >= 50)
                    {
                        items.Add("...");
                        break;
                    }

                    items.Add(FormatLookupValue(item));
                    count++;
                }

                return string.Join("; ", items);
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string TrimForLookup(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= maxLength)
                return value;

            return value.Substring(0, maxLength) + "...";
        }

        //======================================================================
        //команды для обработки сложных свойств/методов элемента 
        //======================================================================

        private static void AddSpecialLookupProperties(Document doc, Element elem, List<object> specialProperties, int maxValueLength)
        {
            if (elem is IndependentTag tag)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetTaggedLocalElementIds",
                        source = "IndependentTag.GetTaggedLocalElementIds()",
                        value = tag.GetTaggedLocalElementIds()
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetTaggedLocalElementIds", "IndependentTag.GetTaggedLocalElementIds()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetTaggedElementIds",
                        source = "IndependentTag.GetTaggedElementIds()",
                        value = tag.GetTaggedElementIds()
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetTaggedElementIds", "IndependentTag.GetTaggedElementIds()", ex));
                }
            }

            if (elem is Dimension dimension)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "References",
                        source = "Dimension.References",
                        value = FormatReferenceArray(doc, dimension.References, maxValueLength)
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("References", "Dimension.References", ex));
                }
            }

            if (elem is View3D view3D)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetSectionBox",
                        source = "View3D.GetSectionBox()",
                        value = FormatBoundingBoxXyz(view3D.GetSectionBox())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetSectionBox", "View3D.GetSectionBox()", ex));
                }
            }

            if (elem is ViewSheet viewSheet)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetAllViewports",
                        source = "ViewSheet.GetAllViewports()",
                        value = FormatElementIdCollectionForLookup(doc, viewSheet.GetAllViewports())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetAllViewports", "ViewSheet.GetAllViewports()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetAllPlacedViews",
                        source = "ViewSheet.GetAllPlacedViews()",
                        value = FormatElementIdCollectionForLookup(doc, viewSheet.GetAllPlacedViews())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetAllPlacedViews", "ViewSheet.GetAllPlacedViews()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "Outline",
                        source = "ViewSheet.Outline",
                        value = FormatBoundingBoxUv(viewSheet.Outline)
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("Outline", "ViewSheet.Outline", ex));
                }
            }

            if (elem is Viewport viewport)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetBoxCenter",
                        source = "Viewport.GetBoxCenter()",
                        value = FormatXyz(viewport.GetBoxCenter())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetBoxCenter", "Viewport.GetBoxCenter()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetBoxOutline",
                        source = "Viewport.GetBoxOutline()",
                        value = FormatOutline(viewport.GetBoxOutline())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetBoxOutline", "Viewport.GetBoxOutline()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetLabelOutline",
                        source = "Viewport.GetLabelOutline()",
                        value = FormatOutline(viewport.GetLabelOutline())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetLabelOutline", "Viewport.GetLabelOutline()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetLabelOffset",
                        source = "Viewport.GetLabelOffset()/LabelOffset",
                        value = FormatXyz(GetViewportLabelOffset(viewport))
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetLabelOffset", "Viewport.GetLabelOffset()", ex));
                }

                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetLabelLineLength",
                        source = "Viewport.GetLabelLineLength()/LabelLineLength",
                        value = GetViewportLabelLineLength(viewport)
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetLabelLineLength", "Viewport.GetLabelLineLength()", ex));
                }
            }

            if (elem is Room room)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetBoundarySegments",
                        source = "Room.GetBoundarySegments()",
                        value = FormatRoomBoundarySegments(doc, room)
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetBoundarySegments", "Room.GetBoundarySegments()", ex));
                }
            }

            if (elem is Autodesk.Revit.DB.Group group)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetMemberIds",
                        source = "Group.GetMemberIds()",
                        value = FormatElementIdCollectionForLookup(doc, group.GetMemberIds())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetMemberIds", "Group.GetMemberIds()", ex));
                }
            }

            if (elem is AssemblyInstance assemblyInstance)
            {
                try
                {
                    specialProperties.Add(new
                    {
                        name = "GetMemberIds",
                        source = "AssemblyInstance.GetMemberIds()",
                        value = FormatElementIdCollectionForLookup(doc, assemblyInstance.GetMemberIds())
                    });
                }
                catch (Exception ex)
                {
                    specialProperties.Add(CreateSpecialPropertyError("GetMemberIds", "AssemblyInstance.GetMemberIds()", ex));
                }

            }



        }

        //описывает какой метод не удалось вызыватьи передает инфу в результат
        private static object CreateSpecialPropertyError(string name, string source, Exception ex)
        {
            return new
            {
                name = name,
                source = source,
                error = ex.Message,
                exception_type = ex.GetType().FullName
            };
        }

        private static List<object> FormatReferenceArray(Document doc, ReferenceArray references, int maxValueLength)
        {
            var result = new List<object>();
            if (references == null)
                return result;

            foreach (Reference reference in references)
            {
                result.Add(reference.ElementId);

            }

            return result;
        }

        private static List<object> FormatRoomBoundarySegments(Document doc, Room room)
        {
            var result = new List<object>();
            if (room == null)
                return result;

            IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(new SpatialElementBoundaryOptions());
            if (boundaries == null)
                return result;

            for (int loopIndex = 0; loopIndex < boundaries.Count; loopIndex++)
            {
                IList<BoundarySegment> loop = boundaries[loopIndex];
                var segments = new List<object>();

                if (loop != null)
                {
                    for (int segmentIndex = 0; segmentIndex < loop.Count; segmentIndex++)
                    {
                        BoundarySegment segment = loop[segmentIndex];
                        segments.Add(FormatBoundarySegment(doc, segment, segmentIndex));
                    }
                }

                result.Add(new
                {
                    loop_index = loopIndex,
                    segment_count = segments.Count,
                    segments = segments
                });
            }

            return result;
        }

        private static object FormatBoundarySegment(Document doc, BoundarySegment segment, int segmentIndex)
        {
            if (segment == null)
            {
                return new
                {
                    segment_index = segmentIndex,
                    is_null = true
                };
            }

            Curve curve = null;
            try
            {
                curve = segment.GetCurve();
            }
            catch
            {
            }

            ElementId elementId = segment.ElementId;
            return new
            {
                segment_index = segmentIndex,
                boundary_element = FormatSingleElementIdForLookup(doc, elementId),
                curve_type = curve != null ? curve.GetType().Name : null,
                length = curve != null ? curve.Length : (double?)null,
                start_point = GetCurveEndPoint(curve, 0),
                end_point = GetCurveEndPoint(curve, 1)
            };
        }

        private static object GetCurveEndPoint(Curve curve, int index)
        {
            if (curve == null)
                return null;

            try
            {
                return FormatXyz(curve.GetEndPoint(index));
            }
            catch
            {
                return null;
            }
        }

        private static List<object> FormatElementIdCollectionForLookup(Document doc, IEnumerable<ElementId> elementIds)
        {
            var result = new List<object>();
            if (elementIds == null)
                return result;

            foreach (ElementId elementId in elementIds)
            {
                result.Add(FormatSingleElementIdForLookup(doc, elementId));
            }

            return result;
        }

        private static object FormatSingleElementIdForLookup(Document doc, ElementId elementId)
        {
            Element element = elementId == null || elementId == ElementId.InvalidElementId
                ? null
                : doc.GetElement(elementId);

            return new
            {
                id = elementId == null || elementId == ElementId.InvalidElementId ? (int?)null : IDHelper.ElIdInt(elementId),
                is_valid = elementId != null && elementId != ElementId.InvalidElementId,
                exists_in_current_document = element != null,
                name = element != null ? element.Name : null,
                category = element != null && element.Category != null ? element.Category.Name : null,
                class_name = element != null ? element.GetType().FullName : null
            };
        }

        private static object FormatBoundingBoxXyz(BoundingBoxXYZ boundingBox)
        {
            if (boundingBox == null)
                return null;

            return new
            {
                min = FormatXyz(boundingBox.Min),
                max = FormatXyz(boundingBox.Max),
                transform = FormatLookupValue(boundingBox.Transform),
                enabled = boundingBox.Enabled
            };
        }

        private static object FormatBoundingBoxUv(BoundingBoxUV boundingBox)
        {
            if (boundingBox == null)
                return null;

            return new
            {
                min = FormatUv(boundingBox.Min),
                max = FormatUv(boundingBox.Max)
            };
        }

        private static object FormatOutline(Outline outline)
        {
            if (outline == null)
                return null;

            return new
            {
                minimum_point = FormatXyz(outline.MinimumPoint),
                maximum_point = FormatXyz(outline.MaximumPoint)
            };
        }

        private static XYZ GetViewportLabelOffset(Viewport viewport)
        {
            object value = InvokeViewportMember(viewport, "GetLabelOffset", "LabelOffset");
            return value as XYZ;
        }

        private static double? GetViewportLabelLineLength(Viewport viewport)
        {
            object value = InvokeViewportMember(viewport, "GetLabelLineLength", "LabelLineLength");
            if (value == null)
                return null;

            try
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
            }
        }

        private static object InvokeViewportMember(Viewport viewport, string methodName, string propertyName)
        {
            if (viewport == null)
                return null;

            Type viewportType = viewport.GetType();
            MethodInfo method = viewportType.GetMethod(methodName, BindingFlags.Public | BindingFlags.Instance, null, Type.EmptyTypes, null);
            if (method != null)
                return method.Invoke(viewport, null);

            PropertyInfo property = viewportType.GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            if (property != null && property.CanRead)
                return property.GetValue(viewport, null);

            return null;
        }

        private static object FormatXyz(XYZ point)
        {
            if (point == null)
                return null;

            return new
            {
                x = point.X,
                y = point.Y,
                z = point.Z
            };
        }

        private static object FormatUv(UV point)
        {
            if (point == null)
                return null;

            return new
            {
                u = point.U,
                v = point.V
            };
        }






        #endregion

        #region 17_get_location_for_element_ids   

        public class PointDto
        {
            public double X { get; set; }
            public double Y { get; set; }
            public double Z { get; set; }
        }

        public class CurveDto
        {
            public PointDto Start { get; set; }
            public PointDto End { get; set; }
        }

        public class ElementLocationInfo
        {
            public PointDto Point { get; set; }   // если точка
            public CurveDto Curve { get; set; }   // если кривая
        }

        public static object GetLocationForElementIds(Document doc, List<int> elementIds)
        {
            var locations = new Dictionary<int, ElementLocationInfo>();

            if (elementIds == null || elementIds.Count == 0)
                return new { locations };

            foreach (int eidInt in elementIds.Distinct())
            {
                var elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null)
                    continue;

                var loc = elem.Location;

                if (loc == null)
                    continue;

                var info = new ElementLocationInfo();

                if (loc is LocationPoint lp)
                {
                    var p = lp.Point;
                    info.Point = new PointDto { X = p.X, Y = p.Y, Z = p.Z };
                }
                else if (loc is LocationCurve lc)
                {
                    var c = lc.Curve;
                    if (c == null) continue;
                    try
                    {
                        if (c is Line || c is Arc || c is NurbSpline)
                        {
                            var s = c.GetEndPoint(0);
                            var e = c.GetEndPoint(1);
                            // ... создание CurveDto
                            info.Curve = new CurveDto
                            {
                                Start = new PointDto { X = s.X, Y = s.Y, Z = s.Z },
                                End = new PointDto { X = e.X, Y = e.Y, Z = e.Z }
                            };
                        }
                        else
                        {
                            // Для других типов кривых использовать аппроксимацию
                            // или пропускать
                            continue;
                        }

                    }
                    catch
                    {
                        // Логировать ошибку и пропустить элемент
                        continue;
                    }
                }
                if (info.Point != null || info.Curve != null)
                    locations[eidInt] = info;
            }

            return new
            {
                locations
            };
        }


        #endregion

        #region 18_get_boundingboxes_for_element_ids   

        public class BoundingBoxDto
        {
            public PointDto min { get; set; }
            public PointDto max { get; set; }
        }

        public static object GetBoundingBoxesForElementIds(Document doc, List<int> elementIds, int? idSheet = null)
        {
            var boundingBoxes = new Dictionary<int, BoundingBoxDto>();

            if (elementIds == null || elementIds.Count == 0)
                return new { bounding_boxes = boundingBoxes };

            foreach (int eidInt in elementIds.Distinct())
            {
                var elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null) continue;

                var view = idSheet.HasValue
                    ? doc.GetElement(IDHelper.ToElementId(idSheet.Value)) as View
                    : doc.ActiveView;

                var bbox = elem.get_BoundingBox(view);
                if (bbox != null)
                {
                    boundingBoxes[eidInt] = new BoundingBoxDto
                    {
                        min = new PointDto
                        {
                            X = bbox.Min.X,
                            Y = bbox.Min.Y,
                            Z = bbox.Min.Z
                        },
                        max = new PointDto
                        {
                            X = bbox.Max.X,
                            Y = bbox.Max.Y,
                            Z = bbox.Max.Z
                        }
                    };
                }
            }

            return new { bounding_boxes = boundingBoxes };
        }

        #endregion

        #region 19_get_boundary_lines   

        public class BoundaryLineInfo
        {
            public double StartX { get; set; }
            public double StartY { get; set; }
            public double StartZ { get; set; }
            public double EndX { get; set; }
            public double EndY { get; set; }
            public double EndZ { get; set; }
            public double Length { get; set; }
        }

        public class ElementBoundaryInfo
        {
            public int ElementId { get; set; }
            public List<BoundaryLineInfo> Lines { get; set; }
            public int LineCount { get; set; }
            public string Error { get; set; }
        }

        public static object GetBoundaryLines(Document doc, List<int> elementIds)
        {
            var results = new List<ElementBoundaryInfo>();

            if (elementIds == null || elementIds.Count == 0)
                return new { boundaries = results, count = 0 };

            var options = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            foreach (int eid in elementIds.Distinct())
            {
                var elementInfo = new ElementBoundaryInfo
                {
                    ElementId = eid,
                    Lines = new List<BoundaryLineInfo>(),
                    LineCount = 0
                };

                Element elem = doc.GetElement(IDHelper.ToElementId(eid));

                if (elem == null)
                {
                    elementInfo.Error = $"Element {eid} not found";
                    results.Add(elementInfo);
                    continue;
                }

                try
                {
                    // Получаем геометрию элемента
                    GeometryElement geomElem = elem.get_Geometry(options);

                    if (geomElem == null)
                    {
                        elementInfo.Error = $"No geometry found for element {eid}";
                        results.Add(elementInfo);
                        continue;
                    }

                    // Рекурсивный обход геометрии
                    ExtractLinesFromGeometry(geomElem, elementInfo.Lines);

                    elementInfo.LineCount = elementInfo.Lines.Count;

                    if (elementInfo.LineCount == 0)
                    {
                        elementInfo.Error = "No boundary lines found";
                    }
                }
                catch (Exception ex)
                {
                    elementInfo.Error = $"Error getting geometry: {ex.Message}";
                }

                results.Add(elementInfo);
            }

            return new
            {
                boundaries = results,
                count = results.Count,
                totalLines = results.Sum(r => r.LineCount)
            };

        }

        // Рекурсивный метод для извлечения линий из геометрии
        private static void ExtractLinesFromGeometry(GeometryElement geomElem, List<BoundaryLineInfo> lines)
        {
            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Line line)
                {
                    // Прямая линия
                    lines.Add(new BoundaryLineInfo
                    {
                        StartX = line.GetEndPoint(0).X,
                        StartY = line.GetEndPoint(0).Y,
                        StartZ = line.GetEndPoint(0).Z,
                        EndX = line.GetEndPoint(1).X,
                        EndY = line.GetEndPoint(1).Y,
                        EndZ = line.GetEndPoint(1).Z,
                        Length = line.Length
                    });
                }
                else if (geomObj is Arc arc)
                {
                    // Дугу можно аппроксимировать несколькими прямыми или вернуть как дугу
                    // Для простоты - разбиваем на сегменты
                    var tessellated = arc.Tessellate();
                    for (int i = 0; i < tessellated.Count - 1; i++)
                    {
                        var p1 = tessellated[i];
                        var p2 = tessellated[i + 1];
                        lines.Add(new BoundaryLineInfo
                        {
                            StartX = p1.X,
                            StartY = p1.Y,
                            StartZ = p1.Z,
                            EndX = p2.X,
                            EndY = p2.Y,
                            EndZ = p2.Z,
                            Length = p1.DistanceTo(p2)
                        });
                    }
                }
                else if (geomObj is PolyLine polyLine)
                {
                    // Полилиния
                    var points = polyLine.GetCoordinates();
                    for (int i = 0; i < points.Count - 1; i++)
                    {
                        var p1 = points[i];
                        var p2 = points[i + 1];
                        lines.Add(new BoundaryLineInfo
                        {
                            StartX = p1.X,
                            StartY = p1.Y,
                            StartZ = p1.Z,
                            EndX = p2.X,
                            EndY = p2.Y,
                            EndZ = p2.Z,
                            Length = p1.DistanceTo(p2)
                        });
                    }
                }
                else if (geomObj is GeometryInstance geomInst)
                {
                    // Рекурсивно обрабатываем вложенную геометрию (символы семейств)
                    ExtractLinesFromGeometry(geomInst.GetInstanceGeometry(), lines);
                }
                else if (geomObj is Solid solid)
                {
                    // Извлекаем рёбра из твердого тела
                    foreach (Edge edge in solid.Edges)
                    {
                        var curve = edge.AsCurve();
                        if (curve is Line line2)
                        {
                            lines.Add(new BoundaryLineInfo
                            {
                                StartX = line2.GetEndPoint(0).X,
                                StartY = line2.GetEndPoint(0).Y,
                                StartZ = line2.GetEndPoint(0).Z,
                                EndX = line2.GetEndPoint(1).X,
                                EndY = line2.GetEndPoint(1).Y,
                                EndZ = line2.GetEndPoint(1).Z,
                                Length = line2.Length
                            });
                        }
                        else if (curve is Arc arc2)
                        {
                            var tessellated = arc2.Tessellate();
                            for (int i = 0; i < tessellated.Count - 1; i++)
                            {
                                var p1 = tessellated[i];
                                var p2 = tessellated[i + 1];
                                lines.Add(new BoundaryLineInfo
                                {
                                    StartX = p1.X,
                                    StartY = p1.Y,
                                    StartZ = p1.Z,
                                    EndX = p2.X,
                                    EndY = p2.Y,
                                    EndZ = p2.Z,
                                    Length = p1.DistanceTo(p2)
                                });
                            }
                        }
                    }
                }
            }
        }
        #endregion

        #region 19.1_get_room_boundary_lines

        public static object GetRoomBoundaryLines(Document doc, List<int> roomIds)
        {
            var results = new List<ElementBoundaryInfo>();

            foreach (int roomId in roomIds)
            {
                var roomInfo = new ElementBoundaryInfo
                {
                    ElementId = roomId,
                    Lines = new List<BoundaryLineInfo>()
                };

                Room room = doc.GetElement(IDHelper.ToElementId(roomId)) as Room;

                if (room == null)
                {
                    roomInfo.Error = $"Room {roomId} not found";
                    results.Add(roomInfo);
                    continue;
                }

                // Получаем границы помещения
                IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(
                    new SpatialElementBoundaryOptions()
                );

                if (boundaries == null || boundaries.Count == 0)
                {
                    roomInfo.Error = "No boundaries found for room";
                    results.Add(roomInfo);
                    continue;
                }

                foreach (IList<BoundarySegment> boundaryLoop in boundaries)
                {
                    foreach (BoundarySegment segment in boundaryLoop)
                    {
                        Curve curve = segment.GetCurve();

                        if (curve is Line line)
                        {
                            roomInfo.Lines.Add(new BoundaryLineInfo
                            {
                                StartX = line.GetEndPoint(0).X,
                                StartY = line.GetEndPoint(0).Y,
                                StartZ = line.GetEndPoint(0).Z,
                                EndX = line.GetEndPoint(1).X,
                                EndY = line.GetEndPoint(1).Y,
                                EndZ = line.GetEndPoint(1).Z,
                                Length = line.Length
                            });
                        }
                        else if (curve is Arc arc)
                        {
                            // Аппроксимация дуги
                            var tessellated = arc.Tessellate();
                            for (int i = 0; i < tessellated.Count - 1; i++)
                            {
                                var p1 = tessellated[i];
                                var p2 = tessellated[i + 1];
                                roomInfo.Lines.Add(new BoundaryLineInfo
                                {
                                    StartX = p1.X,
                                    StartY = p1.Y,
                                    StartZ = p1.Z,
                                    EndX = p2.X,
                                    EndY = p2.Y,
                                    EndZ = p2.Z,
                                    Length = p1.DistanceTo(p2)
                                });
                            }
                        }
                    }
                }

                roomInfo.LineCount = roomInfo.Lines.Count;
                results.Add(roomInfo);
            }

            return new
            {
                boundaries = results,
                count = results.Count,
                totalLines = results.Sum(r => r.LineCount)
            };
        }

        #endregion

        #region 20_get_host_id_for_element_ids

        public static object GetHostIdForElementIds(Document doc, List<int> elementIds)
        {
            var hostIds = new Dictionary<int, int>();

            if (elementIds == null || elementIds.Count == 0)
                return new { host_ids = hostIds };

            foreach (int eidInt in elementIds.Distinct())
            {
                Element elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null)
                    continue;

                int? hostId = null;

                // FamilyInstance (окна, двери, сантехника) — основной случай
                if (elem is FamilyInstance fi)
                {
                    hostId = fi.Host != null ? IDHelper.ElIdInt(fi.Host.Id) : (int?)null;
                }
                // MEPCurve (трубы, воздуховоды) — HostId
                else if (elem is InsulationLiningBase lining)
                {
                    hostId = lining.HostElementId != null ? IDHelper.ElIdInt(lining.HostElementId) : (int?)null;
                }
                // WallSweeps, InsulationLining — GetHostIds()
                else if (elem is WallSweep wallSweep)
                {
                    var hostIdsList = wallSweep.GetHostIds();
                    if (hostIdsList != null && hostIdsList.Count > 0)
                        hostId = IDHelper.ElIdInt(hostIdsList[0]);
                }
                // AreaReinforcement и подобные
                else if (elem is AreaReinforcement areaReinforcement)
                {
                    ElementId areaHostId = areaReinforcement.GetHostId();
                    hostId = areaHostId != null ? IDHelper.ElIdInt(areaHostId) : (int?)null;
                }

                // Если хост найден — добавляем в результат
                if (hostId.HasValue && hostId.Value > 0)
                {
                    hostIds[eidInt] = hostId.Value;
                }
            }

            return new { host_ids = hostIds };
        }

        #endregion

        #region 21_get_object_classes_from_elementids   

        public static object GetObjectClassesFromElementIds(Document doc, List<int> elementIds)
        {
            var objectClasses = new Dictionary<int, string>();

            if (elementIds == null || elementIds.Count == 0)
                return new { object_classes = objectClasses };

            foreach (int eidInt in elementIds.Distinct())
            {
                Element elem = doc.GetElement(IDHelper.ToElementId(eidInt));
                if (elem == null)
                    continue;

                // Полное имя C# класса Revit API
                string className = elem.GetType().FullName ?? "Unknown";
                objectClasses[eidInt] = className;
            }

            return new { object_classes = objectClasses };
        }

        #endregion

        #region 22_get_material_layers_from_types

        public class MaterialLayerInfo
        {
            public string materialName { get; set; }
            public double thickness { get; set; }      // в футах
            public string function { get; set; }       // Structure / Finish и т.п.
        }

        public static object GetMaterialLayersFromTypes(Document doc, List<int> typeIds)
        {
            var result = new Dictionary<int, List<MaterialLayerInfo>>();

            if (typeIds == null || typeIds.Count == 0)
                return new { material_layers = result };

            foreach (int typeIdInt in typeIds.Distinct())
            {
                var typeElem = doc.GetElement(IDHelper.ToElementId(typeIdInt)) as ElementType;
                if (typeElem == null)
                    continue;

                // Поддерживаем только системные типы
                CompoundStructure cs = null;

                if (typeElem is WallType wallType)
                    cs = wallType.GetCompoundStructure();
                else if (typeElem is FloorType floorType)
                    cs = floorType.GetCompoundStructure();
                else if (typeElem is RoofType roofType)
                    cs = roofType.GetCompoundStructure();
                else if (typeElem is CeilingType ceilingType)
                    cs = ceilingType.GetCompoundStructure();
                else
                    continue; // не системный тип конструкции

                if (cs == null)
                    continue;

                var layers = new List<MaterialLayerInfo>();

                IList<CompoundStructureLayer> csLayers = cs.GetLayers();
                foreach (var layer in csLayers)
                {
                    double thickness = layer.Width; // в футах (internal units)

                    string matName = "<No material>";
                    if (layer.MaterialId != ElementId.InvalidElementId)
                    {
                        Material mat = doc.GetElement(layer.MaterialId) as Material;
                        if (mat != null)
                            matName = mat.Name;
                    }

                    string func = layer.Function.ToString(); // Structure, Finish, Membrane и т.д.

                    layers.Add(new MaterialLayerInfo
                    {
                        materialName = matName,
                        thickness = thickness,
                        function = func
                    });
                }

                result[typeIdInt] = layers;
            }

            return new
            {
                material_layers = result
            };
        }

        #endregion

        #region 23_get_model_file_info

        public static object GetModelFileInfo(Document doc)
        {
            try
            {
                string localPath = doc.PathName;
                var result = new
                {
                    is_workshared = false,
                    current_file = new
                    {
                        path = (string)null,
                        size_mb = (double?)null,
                        exists = false
                    },
                    central_file = new
                    {
                        path = (string)null,
                        size_mb = (double?)null,
                        exists = false
                    },
                    error = (string)null
                };

                // Проверяем, сохранён ли файл
                if (string.IsNullOrEmpty(localPath))
                {
                    return new
                    {
                        is_workshared = false,
                        current_file = new
                        {
                            path = "Файл не сохранён на диске",
                            size_mb = (double?)null,
                            exists = false
                        },
                        central_file = new
                        {
                            path = (string)null,
                            size_mb = (double?)null,
                            exists = false
                        },
                        error = "Документ не был сохранён"
                    };
                }

                // Информация о текущем файле
                bool currentFileExists = System.IO.File.Exists(localPath);
                double currentFileSizeMB = 0;
                if (currentFileExists)
                {
                    var currentFileInfo = new System.IO.FileInfo(localPath);
                    currentFileSizeMB = Math.Round(currentFileInfo.Length / (1024.0 * 1024.0), 2);
                }

                // Проверяем, является ли документ общим (workshared)
                bool isWorkshared = doc.IsWorkshared;

                var currentFileInfo_obj = new
                {
                    path = localPath,
                    size_mb = currentFileExists ? currentFileSizeMB : (double?)null,
                    exists = currentFileExists
                };

                var centralFileInfo_obj = new
                {
                    path = (string)null,
                    size_mb = (double?)null,
                    exists = false
                };

                // Если это локальная копия общего файла, пытаемся получить информацию о центральном файле
                if (isWorkshared && currentFileExists)
                {
                    try
                    {
                        var basicInfo = BasicFileInfo.Extract(localPath);
                        string centralPath = basicInfo.CentralPath;

                        if (!string.IsNullOrEmpty(centralPath))
                        {
                            // Проверяем существование центрального файла
                            bool centralExists = System.IO.File.Exists(centralPath);
                            double centralSizeMB = 0;

                            if (centralExists)
                            {
                                var centralFileInfo = new System.IO.FileInfo(centralPath);
                                centralSizeMB = Math.Round(centralFileInfo.Length / (1024.0 * 1024.0), 2);
                            }

                            centralFileInfo_obj = new
                            {
                                path = centralPath,
                                size_mb = centralExists ? centralSizeMB : (double?)null,
                                exists = centralExists
                            };
                        }
                        else
                        {
                            centralFileInfo_obj = new
                            {
                                path = "Не удалось определить путь к центральному файлу",
                                size_mb = (double?)null,
                                exists = false
                            };
                        }
                    }
                    catch (Exception ex)
                    {
                        centralFileInfo_obj = new
                        {
                            path = $"Ошибка при получении информации: {ex.Message}",
                            size_mb = (double?)null,
                            exists = false
                        };
                    }
                }

                // Формируем результат
                return new
                {
                    is_workshared = isWorkshared,
                    current_file = currentFileInfo_obj,
                    central_file = centralFileInfo_obj,
                    summary = isWorkshared
                        ? $"Это локальная копия. Текущий файл: {currentFileSizeMB} МБ, Центральный файл: {(centralFileInfo_obj.size_mb != null ? centralFileInfo_obj.size_mb + " МБ" : "недоступен")}"
                        : $"Это обычный (не общий) файл. Размер: {currentFileSizeMB} МБ"
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    is_workshared = false,
                    current_file = new
                    {
                        path = (string)null,
                        size_mb = (double?)null,
                        exists = false
                    },
                    central_file = new
                    {
                        path = (string)null,
                        size_mb = (double?)null,
                        exists = false
                    },
                    error = $"Ошибка при получении информации о файле: {ex.Message}"
                };
            }
        }

        #endregion

        #region 24_get_all_project_units
#if R2020
        public static object GetAllProjectUnits(Document doc)
        {
            try
            {
                Units units = doc.GetUnits();
                var result = new List<object>();

                // Ключевые типы единиц, которые чаще всего нужны
                var unitTypesToCheck = new Dictionary<UnitType, string>
                {
                    { UnitType.UT_Length, "Length" },
                    { UnitType.UT_Area, "Area" },
                    { UnitType.UT_Volume, "Volume" },
                    { UnitType.UT_Angle, "Angle" },
                    { UnitType.UT_Mass, "Mass" },
                    { UnitType.UT_Currency, "Cost" },
                    { UnitType.UT_TimeInterval, "Time" }
                };
                //UnitType.UT_TimeInterval
                
                        foreach (var kvp in unitTypesToCheck)
                        {
                            try
                            {
                                FormatOptions formatOpt = units.GetFormatOptions(kvp.Key);
                                if (formatOpt != null)
                                {
                                    string displayUnit = formatOpt.DisplayUnits.ToString();
                                    string symbol = GetUnitSymbol(displayUnit);

                                    result.Add(new
                                    {
                                        unitType = kvp.Value,
                                        unitSymbol = symbol,
                                        displayUnitType = displayUnit.Replace("DUT_", ""),
                                        format = formatOpt.GetType().Name
                                    });
                                }
                            }
                            catch { }
                        }

                        return new
                        {
                            project_units = result,
                            count = result.Count
                        };
                    }
                    catch (Exception ex)
                    {
                        return new { error = ex.Message, project_units = new List<object>() };
                    }
                }

        private static string GetUnitSymbol(string displayUnit)
        {
            // 1. Быстрый поиск в нашем словаре
            var symbols = new Dictionary<string, string>
            {
                { "DUT_METERS", "м" },
                { "DUT_MILLIMETERS", "мм" },
                { "DUT_CENTIMETERS", "см" },
                { "DUT_FEET", "фт" },
                { "DUT_INCHES", "дюйм" },
                { "DUT_DECIMAL_FEET", "фт" },
                { "DUT_SQUARE_METERS", "м²" },
                { "DUT_SQUARE_FEET", "фт²" },
                { "DUT_SQUARE_MILLIMETERS", "мм²" },
                { "DUT_CUBIC_METERS", "м³" },
                { "DUT_CUBIC_FEET", "фт³" },
                { "DUT_DEGREES", "°" },
                { "DUT_RADIANS", "рад" },
                { "DUT_KILOGRAMS", "кг" },
                { "DUT_POUNDS", "фунт" },
                { "DUT_CELSIUS", "°C" },
                { "DUT_FAHRENHEIT", "°F" },
                { "DUT_KELVIN", "K" },
                { "DUT_DOLLARS", "$" },
                { "DUT_EUROS", "€" },
                { "DUT_RUBLES", "₽" },
                { "DUT_SECONDS", "с" },
                { "DUT_MINUTES", "мин" },
                { "DUT_HOURS", "ч" },
                { "DUT_PERCENTAGE", "%" },
                { "DUT_LITERS", "л" },
                { "DUT_GALLONS_US", "гал" }
            };

                    if (symbols.ContainsKey(displayUnit))
                        return symbols[displayUnit];

                    // 2. Пробуем получить локализованное название через LabelUtils
                    try
                    {
                        // Парсим строку обратно в DisplayUnitType
                        if (Enum.TryParse(displayUnit, out DisplayUnitType dut))
                        {
                            // Получаем допустимые символы единиц для этого DisplayUnitType
                            var validSymbols = FormatOptions.GetValidUnitSymbols(dut);
                            if (validSymbols != null && validSymbols.Count > 0)
                            {
                                // Берём первый подходящий символ
                                var firstSymbol = validSymbols.FirstOrDefault(s => s != UnitSymbolType.UST_NONE);
                                if (firstSymbol != UnitSymbolType.UST_NONE)
                                {
                                    string localizedLabel = LabelUtils.GetLabelFor(firstSymbol);
                                    if (!string.IsNullOrEmpty(localizedLabel))
                                        return localizedLabel;
                                }
                            }
                        }
                    }
                    catch { /* Игнорируем ошибки */ }

                    // 3. Запасной вариант: убираем префикс DUT_
                    return displayUnit.Replace("DUT_", "");
                }
#endif
        #endregion

        #region 25_get_all_warnings_in_the_model   

        public class WarningInfo
        {
            public string Description { get; set; }
            public string Severity { get; set; }
            public List<int> ElementIds { get; set; } = new List<int>();
        }

        public static object GetAllWarningsInTheModel(Document doc)
        {
            if (doc == null)
                return new { error = "Document is null", warnings = new List<object>() };


            var warningsList = new List<object>();

            // Получаем все постоянные предупреждения документа
            IList<FailureMessage> warnings = doc.GetWarnings();

            foreach (FailureMessage warning in warnings)
            {
                // Для FailureMessage API немного отличается от FailureMessageAccessor
                IList<ElementId> failingElementIds = warning.GetFailingElements().ToList();

                warningsList.Add(new
                {
                    description = warning.GetDescriptionText(),
                    severity = warning.GetSeverity().ToString(),
                    element_ids = failingElementIds.Select(id => IDHelper.ElIdInt(id)).ToList()
                });
            }

            return new
            {
                warnings = warningsList,
                count = warningsList.Count,
                has_warnings = warningsList.Count > 0
            };

        }

        #endregion

        #region 26_get_all_workset_information 

        public static object GetAllWorksetInformation(Document doc)
        {
            try
            {
                if (!doc.IsWorkshared)
                {
                    return new
                    {
                        error = "Документ не является общим (workshared). Рабочие наборы отсутствуют.",
                        is_workshared = false,
                        worksets = new List<object>(),
                        count = 0
                    };
                }

                var worksetsList = new List<object>();

                // Получаем ВСЕ рабочие наборы одним вызовом
                FilteredWorksetCollector collector = new FilteredWorksetCollector(doc);
                IList<Workset> allWorksets = collector.ToWorksets();

                foreach (Workset workset in allWorksets)
                {
                    worksetsList.Add(new
                    {
                        worksetId = workset.Id.IntegerValue,
                        name = workset.Name ?? "Unnamed",
                        owner = workset.Owner ?? "None",
                        isEditable = workset.IsEditable,
                        kind = workset.Kind.ToString(),
                        isOpen = workset.IsOpen,
                        isDefaultWorkset = workset.IsDefaultWorkset,
                        isVisibleByDefault = workset.IsVisibleByDefault
                    });
                }

                // Подсчитываем количество разных типов
                int userWorksetsCount = allWorksets.Count(w => w.Kind == WorksetKind.UserWorkset);
                int systemWorksetsCount = allWorksets.Count - userWorksetsCount;

                return new
                {
                    is_workshared = true,
                    current_user = doc.Application.Username,
                    worksets = worksetsList,
                    count = worksetsList.Count,
                    user_worksets_count = userWorksetsCount,
                    system_worksets_count = systemWorksetsCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении информации о рабочих наборах: {ex.Message}",
                    is_workshared = false,
                    worksets = new List<object>(),
                    count = 0
                };
            }
        }

        #endregion

        #region 27_get_worksets_from_elementids   

        public static object GetWorksetsFromElementIds(Document doc, List<int> elementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список elementIds пуст или не указан",
                        workset_assignments = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Проверяем, является ли документ общим (workshared)
                if (!doc.IsWorkshared)
                {
                    return new
                    {
                        error = "Документ не является общим (workshared). Рабочие наборы отсутствуют.",
                        is_workshared = false,
                        workset_assignments = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Получаем таблицу рабочих наборов для быстрого доступа
                WorksetTable worksetTable = doc.GetWorksetTable();

                // Результат: словарь elementId -> информация о рабочем наборе
                var assignments = new Dictionary<int, object>();

                int processedCount = 0;
                int notFoundCount = 0;

                foreach (int elementId in elementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(elementId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            assignments[elementId] = new
                            {
                                error = "Элемент не найден в документе",
                                worksetId = (int?)null,
                                worksetName = (string)null,
                                isEditable = (bool?)null
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Получаем ID рабочего набора элемента
                        WorksetId worksetId = element.WorksetId;

                        if (worksetId == null || worksetId.IntegerValue == -1)
                        {
                            assignments[elementId] = new
                            {
                                error = "Элемент не принадлежит рабочему набору",
                                worksetId = (int?)null,
                                worksetName = (string)null,
                                isEditable = (bool?)null
                            };
                            continue;
                        }

                        // Получаем информацию о рабочем наборе
                        Workset workset = worksetTable.GetWorkset(worksetId);

                        if (workset == null)
                        {
                            assignments[elementId] = new
                            {
                                error = "Рабочий набор не найден",
                                worksetId = worksetId.IntegerValue,
                                worksetName = (string)null,
                                isEditable = (bool?)null
                            };
                            continue;
                        }

                        // Сохраняем информацию
                        assignments[elementId] = new
                        {
                            worksetId = workset.Id.IntegerValue,
                            worksetName = workset.Name ?? "Unnamed",
                            isEditable = workset.IsEditable,
                            worksetKind = workset.Kind.ToString(),
                            isOpen = workset.IsOpen
                        };

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        assignments[elementId] = new
                        {
                            error = $"Ошибка при получении информации: {ex.Message}",
                            worksetId = (int?)null,
                            worksetName = (string)null,
                            isEditable = (bool?)null
                        };
                    }
                }

                return new
                {
                    workset_assignments = assignments,
                    count = assignments.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    is_workshared = true
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении рабочих наборов для элементов: {ex.Message}",
                    workset_assignments = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        #endregion

        #region 28_get_worksharing_information_for_element_ids

        public static object GetWorksharingInformationForElementIds(Document doc, List<int> elementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список elementIds пуст или не указан",
                        worksharing_info = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Проверяем, является ли документ общим (workshared)
                if (!doc.IsWorkshared)
                {
                    return new
                    {
                        error = "Документ не является общим (workshared). Информация о совместной работе отсутствует.",
                        is_workshared = false,
                        worksharing_info = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Получаем таблицу рабочих наборов для быстрого доступа
                WorksetTable worksetTable = doc.GetWorksetTable();

                // Результат: словарь elementId -> информация о совместной работе
                var worksharingInfo = new Dictionary<int, object>();

                int processedCount = 0;
                int notFoundCount = 0;

                foreach (int elementId in elementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(elementId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            worksharingInfo[elementId] = new
                            {
                                error = "Элемент не найден в документе",
                                worksetName = (string)null,
                                creator = (string)null,
                                owner = (string)null,
                                lastChangedBy = (string)null
                            };
                            notFoundCount++;
                            continue;
                        }

                        // 1. Получаем информацию о рабочем наборе элемента
                        WorksetId worksetId = element.WorksetId;  // Или WorksharingUtils.GetWorksetId(doc, element.Id)
                        string worksetName = null;

                        if (worksetId != null && worksetId.IntegerValue != -1)
                        {
                            Workset workset = worksetTable.GetWorkset(worksetId);
                            if (workset != null)
                            {
                                worksetName = workset.Name ?? "Unnamed";
                            }
                        }

                        // 2. Получаем расширенную информацию о совместной работе через WorksharingTooltipInfo
                        // Этот метод возвращает создателя, владельца и кто последний изменял [citation:4][citation:5]
                        WorksharingTooltipInfo tooltipInfo = WorksharingUtils.GetWorksharingTooltipInfo(doc, elemId);

                        // 3. (Опционально) Получаем статус редактирования элемента
                        CheckoutStatus checkoutStatus = WorksharingUtils.GetCheckoutStatus(doc, elemId);

                        worksharingInfo[elementId] = new
                        {
                            worksetName = worksetName,
                            creator = tooltipInfo?.Creator ?? "Unknown",           // Кто создал элемент [citation:4]
                            owner = tooltipInfo?.Owner ?? "None",                  // Текущий владелец [citation:4]
                            lastChangedBy = tooltipInfo?.LastChangedBy ?? "Unknown", // Кто последний изменял [citation:4]
                            checkoutStatus = checkoutStatus.ToString(),           // Статус редактирования [citation:6]
                            worksetId = worksetId?.IntegerValue
                        };

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        worksharingInfo[elementId] = new
                        {
                            error = $"Ошибка при получении информации: {ex.Message}",
                            worksetName = (string)null,
                            creator = (string)null,
                            owner = (string)null,
                            lastChangedBy = (string)null
                        };
                    }
                }

                // Получаем имя текущего пользователя для контекста
                string currentUser = doc.Application.Username;

                return new
                {
                    worksharing_info = worksharingInfo,
                    count = worksharingInfo.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    is_workshared = true,
                    current_user = currentUser
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении информации о совместной работе: {ex.Message}",
                    worksharing_info = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        #endregion

        #region 29_get_user_selection_in_revit

        public static object GetUserSelectionInRevit(Document doc, UIDocument uiDoc)
        {
            try
            {
                // Проверяем, есть ли активный UI документ
                if (uiDoc == null)
                {
                    return new
                    {
                        error = "Нет активного UI документа. Пожалуйста, откройте проект Revit.",
                        selected_element_ids = new List<int>(),
                        count = 0
                    };
                }

                // Получаем текущее выделение пользователя
                ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();

                if (selectedIds == null || selectedIds.Count == 0)
                {
                    return new
                    {
                        selected_element_ids = new List<int>(),
                        count = 0,
                        message = "Ничего не выделено. Пожалуйста, выделите элементы в Revit."
                    };
                }

                // Преобразуем ElementId в int
                var elementIds = selectedIds
                    .Select(IDHelper.ElIdInt)
                    .ToList();

                // Дополнительная информация о выделении (опционально)
                var elementsInfo = new List<object>();
                foreach (ElementId id in selectedIds)
                {
                    Element elem = doc.GetElement(id);
                    if (elem != null)
                    {
                        elementsInfo.Add(new
                        {
                            id = IDHelper.ElIdInt(id),
                            name = elem.Name ?? "Unnamed",
                            category = elem.Category?.Name ?? "Unknown"
                        });
                    }
                }

                return new
                {
                    selected_element_ids = elementIds,
                    count = elementIds.Count,
                    elements = elementsInfo,  // Дополнительная информация для контекста
                    message = $"Выделено {elementIds.Count} элементов."
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении выделенных элементов: {ex.Message}",
                    selected_element_ids = new List<int>(),
                    count = 0
                };
            }
        }

        #endregion

        #region 30_set_user_selection_in_revit   

        public static object SetUserSelectionInRevit(Document doc, UIDocument uiDoc, List<int> elementIds)
        {
            try
            {
                // Проверяем, есть ли активный UI документ
                if (uiDoc == null)
                {
                    return new
                    {
                        success = false,
                        error = "Нет активного UI документа. Пожалуйста, откройте проект Revit."
                    };
                }

                // Проверяем входные параметры
                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        success = false,
                        error = "Список elementIds пуст. Нечего выделять.",
                        selected_count = 0
                    };
                }

                // Преобразуем int в ElementId и проверяем существование элементов
                var validElementIds = new List<ElementId>();
                var invalidIds = new List<int>();
                var typeIds = new List<int>();

                foreach (int id in elementIds.Distinct())
                {
                    ElementId elemId = IDHelper.ToElementId(id);
                    Element element = doc.GetElement(elemId);

                    if (element == null)
                    {
                        invalidIds.Add(id);
                        continue;
                    }

                    // Проверяем, что это не тип (ElementType)
                    if (element is ElementType)
                    {
                        typeIds.Add(id);
                        continue;
                    }

                    validElementIds.Add(elemId);
                }

                // Если нет валидных ID
                if (validElementIds.Count == 0)
                {
                    string errorMessage = "Нет валидных element id для выделения.";
                    if (invalidIds.Count > 0)
                        errorMessage += $" Не найдены: {string.Join(", ", invalidIds)}.";
                    if (typeIds.Count > 0)
                        errorMessage += $" Type ids (не поддерживаются): {string.Join(", ", typeIds)}.";

                    return new
                    {
                        success = false,
                        error = errorMessage,
                        invalid_ids = invalidIds,
                        type_ids = typeIds,
                        requested_count = elementIds.Count
                    };
                }

                // Устанавливаем выделение
                uiDoc.Selection.SetElementIds(validElementIds);

                //// Дополнительно: прокручиваем вид, чтобы показать выделенные элементы (опционально)
                //if (validElementIds.Count > 0)
                //{
                //    try
                //    {
                //        uiDoc.ShowElements(validElementIds.First());
                //    }
                //    catch { /* Игнорируем, если не получается показать */ }
                //}

                return new
                {
                    success = true,
                    selected_count = validElementIds.Count,
                    requested_count = elementIds.Count,
                    invalid_count = invalidIds.Count,
                    type_count = typeIds.Count,
                    message = $"Выделено {validElementIds.Count} из {elementIds.Count} запрошенных элементов."
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    success = false,
                    error = $"Ошибка при установке выделения: {ex.Message}"
                };
            }
        }

        #endregion

        #region 31_get_graphic_overrides_for_element_ids_in_view   

        public static object GetGraphicOverridesForElementIdsInView(Document doc, List<int> elementIds, int viewId)
        {
            try
            {
                // Проверяем входные параметры
                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список elementIds пуст или не указан",
                        overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Получаем вид по ID
                ElementId viewElemId = IDHelper.ToElementId(viewId);
                View view = doc.GetElement(viewElemId) as View;

                if (view == null)
                {
                    return new
                    {
                        error = $"Вид с ID {viewId} не найден",
                        overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Проверяем, поддерживает ли вид переопределения
                // Некоторые типы видов (например, пустые листы) не поддерживают Visibility/Graphics Overrides
                if (!SupportsGraphicOverrides(view))
                {
                    return new
                    {
                        error = $"Вид типа '{view.ViewType}' не поддерживает графические переопределения",
                        overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int skippedTypesCount = 0;

                foreach (int elementId in elementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(elementId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            result[elementId] = new
                            {
                                error = "Элемент не найден в документе"
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, что это не тип элемента (ElementType)
                        if (element is ElementType)
                        {
                            result[elementId] = new
                            {
                                error = "Type id (не поддерживается). Только экземпляры элементов."
                            };
                            skippedTypesCount++;
                            continue;
                        }

                        // Получаем переопределения для элемента в указанном виде
                        OverrideGraphicSettings overrides = view.GetElementOverrides(elemId);

                        // Извлекаем информацию о переопределениях
                        var overrideInfo = ExtractOverrideInfo(overrides, element, view);

                        result[elementId] = overrideInfo;
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[elementId] = new
                        {
                            error = $"Ошибка при получении переопределений: {ex.Message}"
                        };
                    }
                }

                return new
                {
                    overrides = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    skipped_types = skippedTypesCount,
                    view_id = viewId,
                    view_name = view.Name,
                    view_type = view.ViewType.ToString()
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении графических переопределений: {ex.Message}",
                    overrides = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        /// <summary>
        /// Проверяет, поддерживает ли вид графические переопределения [citation:5]
        /// </summary>
        private static bool SupportsGraphicOverrides(View view)
        {
            // Виды, которые не поддерживают Visibility/Graphics Overrides
            var unsupportedTypes = new HashSet<ViewType>
            {
                ViewType.Schedule,           // Спецификации
                ViewType.DrawingSheet,       // Листы
                ViewType.Legend,             // Легенды
                ViewType.Report,             // Отчёты
                ViewType.ColumnSchedule,     // Спецификации колонн
                ViewType.PanelSchedule,      // Спецификации панелей
                ViewType.LoadsReport,        // Отчёты по нагрузкам
                ViewType.CostReport,         // Отчёты по стоимости
                ViewType.PresureLossReport,    // Отчёты по давлениям
                ViewType.SystemBrowser,      // Обозреватель систем
                ViewType.Internal,           // Внутренние виды Revit
                ViewType.Undefined,          // Неопределённый тип
                ViewType.Walkthrough         // Проходы (3D-анимация)
            };

            //return !unsupportedTypes.Contains(view.ViewType);

            var supportedTypes = new HashSet<ViewType>
            {
                ViewType.FloorPlan,          // Планы этажей ✅
                ViewType.CeilingPlan,        // Планы потолков ✅
                ViewType.ThreeD,             // 3D-виды ✅
                ViewType.DraftingView,       // Чертёжные виды ✅
                ViewType.Detail,             // Детальные виды ✅
                ViewType.AreaPlan,           // Планы зон ✅
                ViewType.EngineeringPlan,    // Инженерные планы (частично) bkили же план несущих конструкций✅
                ViewType.Elevation,          // Фасады (только для некоторых элементов)
                ViewType.Section             // Разрезы (только для некоторых элементов)
            };

            if (supportedTypes.Contains(view.ViewType))
                return true;

            if (unsupportedTypes.Contains(view.ViewType))
                return false;

            // Для остальных типов (например, Elevation, Section) 
            // пытаемся получить переопределения через try-catch
            try
            {
                // Находим любой элемент в документе
                var collector = new FilteredElementCollector(view.Document)
                    .WhereElementIsNotElementType()
                    .FirstOrDefault();

                if (collector != null)
                {
                    var testOverrides = view.GetElementOverrides(collector.Id);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;

        }



        /// <summary>
        /// Извлекает информацию из OverrideGraphicSettings [citation:1][citation:5]
        /// </summary>
        private static object ExtractOverrideInfo(OverrideGraphicSettings overrides, Element el, View view)
        {
            var projection = new
            {
                // Линии проекции
                line_color = GetColorInfo(overrides.ProjectionLineColor),
                line_pattern_id = overrides.ProjectionLinePatternId != null ? IDHelper.ElIdInt(overrides.ProjectionLinePatternId) : (int?)null,
                line_weight = overrides.ProjectionLineWeight != -1 ? overrides.ProjectionLineWeight : (int?)null,

                // Поверхности
                surface_foreground_pattern_id = overrides.SurfaceForegroundPatternId != null ? IDHelper.ElIdInt(overrides.SurfaceForegroundPatternId) : (int?)null,
                surface_foreground_pattern_color = GetColorInfo(overrides.SurfaceForegroundPatternColor),
                surface_foreground_pattern_visible = overrides.IsSurfaceForegroundPatternVisible,

                surface_background_pattern_id = overrides.SurfaceBackgroundPatternId != null ? IDHelper.ElIdInt(overrides.SurfaceBackgroundPatternId) : (int?)null,
                surface_background_pattern_color = GetColorInfo(overrides.SurfaceBackgroundPatternColor),
                surface_background_pattern_visible = overrides.IsSurfaceBackgroundPatternVisible,
                transparency = overrides.Transparency
            };

            var cut = new
            {
                // Линии разреза
                line_color = GetColorInfo(overrides.CutLineColor),
                line_pattern_id = overrides.CutLinePatternId != null ? IDHelper.ElIdInt(overrides.CutLinePatternId) : (int?)null,
                line_weight = overrides.CutLineWeight != -1 ? overrides.CutLineWeight : (int?)null,

                // Поверхности разреза
                foreground_pattern_id = overrides.CutForegroundPatternId != null ? IDHelper.ElIdInt(overrides.CutForegroundPatternId) : (int?)null,
                foreground_pattern_color = GetColorInfo(overrides.CutForegroundPatternColor),
                foreground_pattern_visible = overrides.IsCutForegroundPatternVisible,

                background_pattern_id = overrides.CutBackgroundPatternId != null ? IDHelper.ElIdInt(overrides.CutBackgroundPatternId) : (int?)null,
                background_pattern_color = GetColorInfo(overrides.CutBackgroundPatternColor),
                background_pattern_visible = overrides.IsCutBackgroundPatternVisible
            };

            return new
            {
                //переопределение проекции
                projection = projection,
                //переопределения в сечении/разрезе
                cut = cut,
                //скрыт ли элемент
                is_hidden = el.IsHidden(view),
                //полутона
                halftone = overrides.Halftone,
                //
                has_overrides = HasAnyOverride(overrides, el, view)
            };
        }

        /// <summary>
        /// Преобразует цвет в читаемый формат
        /// </summary>
        private static object GetColorInfo(Color color)
        {
            if (color == null || !color.IsValid)
                return null;

            return new
            {
                red = color.Red,
                green = color.Green,
                blue = color.Blue,
                is_valid = color.IsValid,
                rgb_string = $"RGB({color.Red}, {color.Green}, {color.Blue})"
            };
        }


        /// <summary>
        /// Проверяет, есть ли какие-либо активные переопределения
        /// </summary>
        private static object HasAnyOverride(OverrideGraphicSettings overrides, Element el, View view)
        {
            if (overrides == null)
                return false;

            // 1. Проверка цветов линий
            bool hasProjectionLineColor = overrides.ProjectionLineColor != null && overrides.ProjectionLineColor.IsValid;
            bool hasCutLineColor = overrides.CutLineColor != null && overrides.CutLineColor.IsValid;

            // 2. Проверка весов линий (-1 означает "не задано")
            bool hasProjectionLineWeight = overrides.ProjectionLineWeight > 0;
            bool hasCutLineWeight = overrides.CutLineWeight > 0;

            // 3. Проверка паттернов линий
            bool hasProjectionLinePattern = overrides.ProjectionLinePatternId != null &&
                                            IDHelper.ElIdInt(overrides.ProjectionLinePatternId) > 0;

            bool hasCutLinePattern = overrides.CutLinePatternId != null &&
                                     IDHelper.ElIdInt(overrides.CutLinePatternId) > 0;

            // 4. Проверка цветов паттернов поверхностей
            bool hasSurfaceForegroundColor = overrides.SurfaceForegroundPatternColor != null &&
                                              overrides.SurfaceForegroundPatternColor.IsValid;
            bool hasSurfaceBackgroundColor = overrides.SurfaceBackgroundPatternColor != null &&
                                              overrides.SurfaceBackgroundPatternColor.IsValid;

            // 5. Проверка цветов паттернов разреза
            bool hasCutForegroundColor = overrides.CutForegroundPatternColor != null &&
                                          overrides.CutForegroundPatternColor.IsValid;
            bool hasCutBackgroundColor = overrides.CutBackgroundPatternColor != null &&
                                          overrides.CutBackgroundPatternColor.IsValid;

            // 6. Проверка ID паттернов поверхностей
            bool hasSurfaceForegroundPattern = overrides.SurfaceForegroundPatternId != null &&
                                               IDHelper.ElIdInt(overrides.SurfaceForegroundPatternId) > 0;
            bool hasSurfaceBackgroundPattern = overrides.SurfaceBackgroundPatternId != null &&
                                               IDHelper.ElIdInt(overrides.SurfaceBackgroundPatternId) > 0;

            // 7. Проверка ID паттернов разреза
            bool hasCutForegroundPattern = overrides.CutForegroundPatternId != null &&
                                           IDHelper.ElIdInt(overrides.CutForegroundPatternId) > 0;
            bool hasCutBackgroundPattern = overrides.CutBackgroundPatternId != null &&
                                           IDHelper.ElIdInt(overrides.CutBackgroundPatternId) > 0;

            // 8. Проверка видимости паттернов
            bool hasSurfaceForegroundVisible = overrides.IsSurfaceForegroundPatternVisible;
            bool hasSurfaceBackgroundVisible = overrides.IsSurfaceBackgroundPatternVisible;
            bool hasCutForegroundVisible = overrides.IsCutForegroundPatternVisible;
            bool hasCutBackgroundVisible = overrides.IsCutBackgroundPatternVisible;

            // 9. Проверка полутона, прозрачности и видимости
            bool hasHalftone = overrides.Halftone;
            bool hasTransparency = overrides.Transparency != 0;
            bool is_hidden = el.IsHidden(view);

            return new
            {
                hasProjectionLineColor,
                hasCutLineColor,
                hasProjectionLineWeight,
                hasCutLineWeight,
                hasProjectionLinePattern,
                hasCutLinePattern,
                hasSurfaceForegroundColor,
                hasSurfaceBackgroundColor,
                hasCutForegroundColor,
                hasCutBackgroundColor,
                hasSurfaceForegroundPattern,
                hasSurfaceBackgroundPattern,
                hasCutForegroundPattern,
                hasCutBackgroundPattern,
                hasSurfaceForegroundVisible,
                hasSurfaceBackgroundVisible,
                hasCutForegroundVisible,
                hasCutBackgroundVisible,
                hasHalftone,
                hasTransparency,
                is_hidden
            };


        }

        #endregion

        #region 32_get_graphic_filters_applied_to_views   

        public static object GetGraphicFiltersAppliedToViews(Document doc, List<int> viewElementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (viewElementIds == null || viewElementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список viewElementIds пуст или не указан",
                        view_filters = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int unsupportedCount = 0;

                foreach (int viewId in viewElementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(viewId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            result[viewId] = new
                            {
                                error = "Вид не найден в документе"
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, является ли элемент видом
                        View view = element as View;
                        if (view == null)
                        {
                            result[viewId] = new
                            {
                                error = "Элемент не является видом. Укажите ID вида или листа."
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, поддерживает ли вид фильтры
                        if (!SupportsFilters(view))
                        {
                            result[viewId] = new
                            {
                                error = $"Вид типа '{view.ViewType}' не поддерживает фильтры. " +
                                        "Фильтры доступны для планов этажей, планов потолков, 3D-видов и чертёжных видов."
                            };
                            unsupportedCount++;
                            continue;
                        }

                        // Получаем ID фильтров, применённых к виду [citation:1][citation:4]
                        ICollection<ElementId> filterIds = view.GetFilters();

                        var filtersList = new List<object>();

                        foreach (ElementId filterId in filterIds)
                        {
                            // Получаем объект фильтра
                            ParameterFilterElement filter = doc.GetElement(filterId) as ParameterFilterElement;

                            if (filter == null)
                                continue;

                            // Получаем категории, к которым применяется фильтр [citation:5]
                            ICollection<ElementId> categoryIds = filter.GetCategories();
                            var categoryIdList = categoryIds?.Select(IDHelper.ElIdInt).ToList() ?? new List<int>();

                            // Получаем видимость фильтра в виде [citation:7]
                            bool isVisible = view.GetFilterVisibility(filterId);

                            // Получаем информацию о правилах фильтра (опционально)
                            var rulesInfo = GetFilterRulesInfo(filter);

                            filtersList.Add(new
                            {
                                filterId = IDHelper.ElIdInt(filterId),
                                filterName = filter.Name ?? "Unnamed",
                                categories = categoryIdList,
                                isVisible = isVisible,
                                hasRules = rulesInfo.hasRules,
                                ruleParameters = rulesInfo.parameters
                            });
                        }

                        result[viewId] = new
                        {
                            filters = filtersList,
                            count = filtersList.Count,
                            // ordered_filter_ids = orderedFilterIds,
                            supports_filters = true
                        };

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[viewId] = new
                        {
                            error = $"Ошибка при получении фильтров: {ex.Message}"
                        };
                    }
                }

                return new
                {
                    view_filters = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    unsupported_views = unsupportedCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении фильтров видов: {ex.Message}",
                    view_filters = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        /// <summary>
        /// Проверяет, поддерживает ли вид фильтры
        /// </summary>
        private static bool SupportsFilters(View view)
        {
            // Виды, которые поддерживают фильтры в VG
            var supportedTypes = new HashSet<ViewType>
            {
                ViewType.FloorPlan,          // Планы этажей
                ViewType.CeilingPlan,        // Планы потолков
                ViewType.ThreeD,             // 3D-виды
                ViewType.DraftingView,       // Чертёжные виды
                ViewType.Detail,             // Детальные виды
                ViewType.AreaPlan,           // Планы зон
                ViewType.EngineeringPlan,    // Инженерные планы
                ViewType.Section,            // Разрезы
                ViewType.Elevation           // Фасады
            };

            return supportedTypes.Contains(view.ViewType);
        }

        /// <summary>
        /// Получает информацию о правилах фильтра
        /// </summary>
        private static (bool hasRules, List<string> parameters) GetFilterRulesInfo(ParameterFilterElement filter)
        {
            var parameters = new List<string>();
            bool hasRules = false;

            try
            {
                // Получаем ElementFilter, представляющий правила фильтра [citation:5]
                ElementFilter elementFilter = filter.GetElementFilter();

                if (elementFilter != null)
                {
                    hasRules = true;

                    // Если это ParameterFilter, можно получить параметры
                    if (elementFilter is ElementParameterFilter paramFilter)
                    {
                        var rules = paramFilter.GetRules();
                        foreach (var rule in rules)
                        {
                            // Пытаемся получить имя параметра
                            if (rule is FilterRule filterRule)
                            {
                                // Через ParameterId можно получить имя параметра
                                //var paramId = filterRule.ParameterId;
                                var paramId = filterRule.GetRuleParameter();
                                parameters.Add(paramId.ToString());
                            }
                        }
                    }
                }
            }
            catch
            {
                // Игнорируем ошибки при получении правил
            }

            return (hasRules, parameters.Distinct().Take(10).ToList());
        }

        #endregion

        #region 32.1_get_all_parameter_filters_in_model

        /// <summary>
        /// Возвращает все фильтры видов (ParameterFilterElement) в модели Revit
        /// </summary>
        public static object GetAllParameterFiltersInModel(Document doc)
        {
            try
            {
                // Собираем все фильтры видов в документе
                var allFilters = new FilteredElementCollector(doc)
                    .OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .ToList();

                if (allFilters.Count == 0)
                {
                    return new
                    {
                        filters = new List<object>(),
                        count = 0,
                        message = "В модели не найдено ни одного фильтра видов."
                    };
                }

                var result = new List<object>();
                int processedCount = 0;
                int errorCount = 0;

                foreach (ParameterFilterElement filter in allFilters)
                {
                    try
                    {
                        // Получаем категории, к которым применяется фильтр
                        ICollection<ElementId> categoryIds = filter.GetCategories();
                        var categoryIdList = categoryIds?.Select(IDHelper.ElIdInt).ToList() ?? new List<int>();

                        // Получаем информацию о правилах фильтра
                        var rulesInfo = GetFilterRulesInfo(filter);

                        result.Add(new
                        {
                            filterId = IDHelper.ElIdInt(filter.Id),
                            filterName = filter.Name ?? "Unnamed",
                            categories = categoryIdList,
                            hasRules = rulesInfo.hasRules,
                            ruleParameters = rulesInfo.parameters  // Список ID параметров в правилах
                        });

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result.Add(new
                        {
                            filterId = IDHelper.ElIdInt(filter.Id),
                            filterName = filter.Name ?? "Unnamed",
                            error = $"Ошибка при получении информации: {ex.Message}"
                        });
                        errorCount++;
                    }
                }

                // Сортируем по имени для удобства
                var sortedResult = result.OrderBy(f =>
                {
                    var nameProp = f.GetType().GetProperty("filterName");
                    return nameProp?.GetValue(f)?.ToString() ?? "";
                }).ToList();

                return new
                {
                    filters = sortedResult,
                    count = sortedResult.Count,
                    processed_successfully = processedCount,
                    errors = errorCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении фильтров: {ex.Message}",
                    filters = new List<object>(),
                    count = 0
                };
            }
        }

        ///// <summary>
        ///// Получает информацию о правилах фильтра
        ///// </summary>
        //private static (bool hasRules, List<string> parameters) GetFilterRulesInfo(ParameterFilterElement filter)
        //{
        //    var parameters = new List<string>();
        //    bool hasRules = false;

        //    try
        //    {
        //        // Получаем ElementFilter, представляющий правила фильтра
        //        ElementFilter elementFilter = filter.GetElementFilter();

        //        if (elementFilter != null)
        //        {
        //            hasRules = true;

        //            // Если это ElementParameterFilter, можно получить параметры
        //            if (elementFilter is ElementParameterFilter paramFilter)
        //            {
        //                var rules = paramFilter.GetRules();
        //                foreach (var rule in rules)
        //                {
        //                    if (rule is FilterRule filterRule)
        //                    {
        //                        try
        //                        {
        //                            // Получаем ID параметра, участвующего в правиле
        //                            ElementId paramId = filterRule.GetRuleParameter();
        //                            if (paramId != null && paramId.IntegerValue != -1)
        //                            {
        //                                parameters.Add(paramId.IntegerValue.ToString());
        //                            }
        //                        }
        //                        catch { }
        //                    }
        //                }
        //            }
        //            // Для логических фильтров (AND/OR) можно рекурсивно обойти
        //            else if (elementFilter is LogicalFilter logicalFilter)
        //            {
        //                // Опционально: можно рекурсивно обойти вложенные фильтры
        //                // Но для простоты оставим как есть
        //            }
        //        }
        //    }
        //    catch
        //    {
        //        // Игнорируем ошибки при получении правил
        //    }

        //    return (hasRules, parameters.Distinct().Take(20).ToList());
        //}

        #endregion

        #region 33_get_graphic_overrides_view_filters 

        public static object GetGraphicOverridesViewFilters(Document doc, List<int> filterIds, int viewId)
        {
            try
            {
                // Проверяем входные параметры
                if (filterIds == null || filterIds.Count == 0)
                {
                    return new
                    {
                        error = "Список filterIds пуст или не указан",
                        filter_overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Получаем вид по ID
                ElementId viewElemId = IDHelper.ToElementId(viewId);
                View view = doc.GetElement(viewElemId) as View;

                if (view == null)
                {
                    return new
                    {
                        error = $"Вид с ID {viewId} не найден",
                        filter_overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                // Проверяем, поддерживает ли вид фильтры
                if (!SupportsFilters(view))
                {
                    return new
                    {
                        error = $"Вид типа '{view.ViewType}' не поддерживает фильтры. " +
                                "Фильтры доступны для планов этажей, планов потолков, 3D-видов и чертёжных видов.",
                        filter_overrides = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int notAppliedCount = 0;

                foreach (int filterId in filterIds.Distinct())
                {
                    try
                    {
                        ElementId filterElemId = IDHelper.ToElementId(filterId);

                        // Проверяем, существует ли фильтр в документе
                        ParameterFilterElement filter = doc.GetElement(filterElemId) as ParameterFilterElement;
                        if (filter == null)
                        {
                            result[filterId] = new
                            {
                                error = "Фильтр не найден в документе"
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Получаем ID всех фильтров, применённых к виду [citation:1][citation:4]
                        ICollection<ElementId> appliedFilters = view.GetFilters();

                        // Проверяем, применён ли данный фильтр к виду
                        if (!appliedFilters.Contains(filterElemId))
                        {
                            result[filterId] = new
                            {
                                error = "Фильтр не применён к данному виду",
                                isApplied = false
                            };
                            notAppliedCount++;
                            continue;
                        }

                        // Получаем графические переопределения для фильтра в данном виде [citation:3][citation:7][citation:9]
                        OverrideGraphicSettings overrides = view.GetFilterOverrides(filterElemId);

                        // Извлекаем информацию о переопределениях
                        var overrideInfo = ExtractFilterOverrideInfo(overrides);

                        result[filterId] = overrideInfo;
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[filterId] = new
                        {
                            error = $"Ошибка при получении переопределений: {ex.Message}"
                        };
                    }
                }

                return new
                {
                    filter_overrides = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    not_applied = notAppliedCount,
                    view_id = viewId,
                    view_name = view.Name,
                    view_type = view.ViewType.ToString()
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении графических переопределений фильтров: {ex.Message}",
                    filter_overrides = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }


        /// <summary>
        /// Извлекает информацию о переопределениях фильтра из OverrideGraphicSettings
        /// </summary>
        private static object ExtractFilterOverrideInfo(OverrideGraphicSettings overrides)
        {



            // Настройки проекции (линии и поверхности)
            var projection = new
            {
                // Линии проекции
                line_color = GetColorInfo(overrides.ProjectionLineColor),
                line_pattern_id = overrides.ProjectionLinePatternId != null ? IDHelper.ElIdInt(overrides.ProjectionLinePatternId) : (int?)null,
                line_weight = overrides.ProjectionLineWeight != -1 ? overrides.ProjectionLineWeight : (int?)null,

                // Заливка поверхности (передний план)
                surface_foreground_pattern_id = overrides.SurfaceForegroundPatternId != null ? IDHelper.ElIdInt(overrides.SurfaceForegroundPatternId) : (int?)null,
                surface_foreground_pattern_color = GetColorInfo(overrides.SurfaceForegroundPatternColor),
                surface_foreground_pattern_visible = overrides.IsSurfaceForegroundPatternVisible,

                // Заливка поверхности (задний план)
                surface_background_pattern_id = overrides.SurfaceBackgroundPatternId != null ? IDHelper.ElIdInt(overrides.SurfaceBackgroundPatternId) : (int?)null,
                surface_background_pattern_color = GetColorInfo(overrides.SurfaceBackgroundPatternColor),
                surface_background_pattern_visible = overrides.IsSurfaceBackgroundPatternVisible,

                // Прозрачность
                transparency = overrides.Transparency
            };

            // Настройки разреза
            var cut = new
            {
                // Линии разреза
                line_color = GetColorInfo(overrides.CutLineColor),
                line_pattern_id = overrides.CutLinePatternId != null ? IDHelper.ElIdInt(overrides.CutLinePatternId) : (int?)null,
                line_weight = overrides.CutLineWeight != -1 ? overrides.CutLineWeight : (int?)null,

                // Заливка разреза (передний план)
                foreground_pattern_id = overrides.CutForegroundPatternId != null ? IDHelper.ElIdInt(overrides.CutForegroundPatternId) : (int?)null,
                foreground_pattern_color = GetColorInfo(overrides.CutForegroundPatternColor),
                foreground_pattern_visible = overrides.IsCutForegroundPatternVisible,

                // Заливка разреза (задний план)
                background_pattern_id = overrides.CutBackgroundPatternId != null ? IDHelper.ElIdInt(overrides.CutBackgroundPatternId) : (int?)null,
                background_pattern_color = GetColorInfo(overrides.CutBackgroundPatternColor),
                background_pattern_visible = overrides.IsCutBackgroundPatternVisible
            };

            return new
            {
                projection = projection,
                cut = cut,
                halftone = overrides.Halftone,
                has_overrides = HasAnyFilterOverride(overrides)
            };
        }

        /// <summary>
        /// Проверяет, есть ли активные переопределения у фильтра
        /// </summary>
        private static object HasAnyFilterOverride(OverrideGraphicSettings overrides)
        {
            if (overrides == null)
                return false;

            // 1. Проверка цветов линий
            bool hasProjectionLineColor = overrides.ProjectionLineColor != null && overrides.ProjectionLineColor.IsValid;
            bool hasCutLineColor = overrides.CutLineColor != null && overrides.CutLineColor.IsValid;

            // 2. Проверка весов линий (-1 означает "не задано")
            bool hasProjectionLineWeight = overrides.ProjectionLineWeight > 0;
            bool hasCutLineWeight = overrides.CutLineWeight > 0;

            // 3. Проверка паттернов линий
            bool hasProjectionLinePattern = overrides.ProjectionLinePatternId != null &&
                                            IDHelper.ElIdInt(overrides.ProjectionLinePatternId) > 0;
            bool hasCutLinePattern = overrides.CutLinePatternId != null &&
                                     IDHelper.ElIdInt(overrides.CutLinePatternId) > 0;

            // 4. Проверка цветов паттернов поверхностей
            bool hasSurfaceForegroundColor = overrides.SurfaceForegroundPatternColor != null &&
                                              overrides.SurfaceForegroundPatternColor.IsValid;
            bool hasSurfaceBackgroundColor = overrides.SurfaceBackgroundPatternColor != null &&
                                              overrides.SurfaceBackgroundPatternColor.IsValid;

            // 5. Проверка цветов паттернов разреза
            bool hasCutForegroundColor = overrides.CutForegroundPatternColor != null &&
                                          overrides.CutForegroundPatternColor.IsValid;
            bool hasCutBackgroundColor = overrides.CutBackgroundPatternColor != null &&
                                          overrides.CutBackgroundPatternColor.IsValid;

            // 6. Проверка ID паттернов поверхностей
            bool hasSurfaceForegroundPattern = overrides.SurfaceForegroundPatternId != null &&
                                               IDHelper.ElIdInt(overrides.SurfaceForegroundPatternId) > 0;
            bool hasSurfaceBackgroundPattern = overrides.SurfaceBackgroundPatternId != null &&
                                               IDHelper.ElIdInt(overrides.SurfaceBackgroundPatternId) > 0;

            // 7. Проверка ID паттернов разреза
            bool hasCutForegroundPattern = overrides.CutForegroundPatternId != null &&
                                           IDHelper.ElIdInt(overrides.CutForegroundPatternId) > 0;
            bool hasCutBackgroundPattern = overrides.CutBackgroundPatternId != null &&
                                           IDHelper.ElIdInt(overrides.CutBackgroundPatternId) > 0;

            // 8. Проверка видимости паттернов (если видимость вылкючена, значит переопределение задано)
            bool hasSurfaceForegroundVisible = !overrides.IsSurfaceForegroundPatternVisible;
            bool hasSurfaceBackgroundVisible = !overrides.IsSurfaceBackgroundPatternVisible;
            bool hasCutForegroundVisible = !overrides.IsCutForegroundPatternVisible;
            bool hasCutBackgroundVisible = !overrides.IsCutBackgroundPatternVisible;

            // 9. Проверка полутона, прозрачности
            bool hasHalftone = overrides.Halftone;
            bool hasTransparency = overrides.Transparency != 0;


            return new
            {
                hasProjectionLineColor,
                hasCutLineColor,
                hasProjectionLineWeight,
                hasCutLineWeight,
                hasProjectionLinePattern,
                hasCutLinePattern,
                hasSurfaceForegroundColor,
                hasSurfaceBackgroundColor,
                hasCutForegroundColor,
                hasCutBackgroundColor,
                hasSurfaceForegroundPattern,
                hasSurfaceBackgroundPattern,
                hasCutForegroundPattern,
                hasCutBackgroundPattern,
                hasSurfaceForegroundVisible,
                hasSurfaceBackgroundVisible,
                hasCutForegroundVisible,
                hasCutBackgroundVisible,
                hasHalftone,
                hasTransparency
            };

        }

        #endregion

        #region 33.1_get_category_visibility_overrides_in_view

        /// <summary>
        /// Возвращает переопределения видимости по категориям на указанном виде
        /// </summary>
        public static object GetCategoryVisibilityOverridesInView(Document doc, int viewId)
        {
            //список переопределений по категориям
            var result = new Dictionary<int, object>();
            int processedCount = 0;

            ElementId viewElemId = IDHelper.ToElementId(viewId);
            View view = doc.GetElement(viewElemId) as View;


            try
            {
                if (view == null)
                {
                    return new
                    {
                        error = $"Вид с ID {viewId} не найден",
                        category_overrides = new List<object>(),
                        count = 0
                    };
                }

                // Проверяем, поддерживает ли вид переопределения категорий
                if (!SupportsCategoryOverrides(view))
                {
                    return new
                    {
                        error = $"Вид типа '{view.ViewType}' не поддерживает переопределения видимости по категориям. " +
                                "Поддерживаются: планы этажей, планы потолков, 3D-виды, чертёжные виды, фасады, разрезы.",
                        category_overrides = new List<object>(),
                        count = 0
                    };
                }



                // Получаем все категории, доступные в документе
                var allCategories = GetCategories(doc);


                foreach (var cat in allCategories)
                {
                    ElementId catId = cat.Id;

                    // Получаем настройки видимости/графики для вида
                    OverrideGraphicSettings overrides = view.GetCategoryOverrides(catId);

                    // Извлекаем информацию о переопределениях
                    var overrideInfo = ExtractFilterOverrideInfo(overrides);

                    bool isHidden = false;
                    try
                    {
                        isHidden = view.GetCategoryHidden(catId);
                    }
                    catch
                    {
                        isHidden = false; // или true, в зависимости от логики
                    }


                    result[IDHelper.ElIdInt(catId)] = new
                    {
                        category_id = IDHelper.ElIdInt(catId),
                        category_name = cat.Name ?? "Unnamed",
                        is_hidden = isHidden,
                        overrides = overrideInfo
                    };
                    processedCount++;

                }
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении переопределений категорий: {ex.Message}",
                    category_overrides = new List<object>(),
                    count = 0
                };
            }


            return new
            {
                categories_overrides = result,
                count = result.Count,
                processed_successfully = processedCount,
                view_id = viewId,
                view_name = view.Name,
                view_type = view.ViewType.ToString()
            };

        }

        /// <summary>
        /// Проверяет, поддерживает ли вид переопределения категорий
        /// </summary>
        private static bool SupportsCategoryOverrides(View view)
        {
            // Виды, которые поддерживают переопределения категорий
            var supportedTypes = new HashSet<ViewType>
            {
                ViewType.FloorPlan,          // Планы этажей
                ViewType.CeilingPlan,        // Планы потолков
                ViewType.ThreeD,             // 3D-виды
                ViewType.DraftingView,       // Чертёжные виды
                ViewType.Detail,             // Детальные виды
                ViewType.AreaPlan,           // Планы зон
                ViewType.EngineeringPlan,    // Инженерные планы
                ViewType.Section,            // Разрезы
                ViewType.Elevation           // Фасады
            };

            return supportedTypes.Contains(view.ViewType);
        }

        /// <summary>
        /// Возвращает все категории в документе
        /// </summary>
        private static List<Category> GetCategories(Document doc)
        {
            var categories = new List<Category>();
            Categories settingsCategories = doc.Settings.Categories;

            foreach (Category cat in settingsCategories)
            {
                // Пропускаем категории, которые не имеют видимых элементов
                if (cat == null) continue;

                categories.Add(cat);
            }

            return categories.OrderBy(c => c.Name).ToList();
        }
        #endregion

        #region 33.2_get_workset_visibility_in_view

        /// <summary>
        /// Возвращает информацию о видимости рабочих наборов на указанном виде
        /// </summary>
        public static object GetWorksetVisibilityInView(Document doc, int viewId)
        {
            try
            {
                // Получаем вид по ID
                ElementId viewElemId = IDHelper.ToElementId(viewId);
                View view = doc.GetElement(viewElemId) as View;

                if (view == null)
                {
                    return new
                    {
                        error = $"Вид с ID {viewId} не найден",
                        workset_visibility = new List<object>(),
                        count = 0
                    };
                }

                // Проверяем, является ли документ общим (workshared)
                if (!doc.IsWorkshared)
                {
                    return new
                    {
                        error = "Документ не является общим (workshared). Рабочие наборы отсутствуют.",
                        workset_visibility = new List<object>(),
                        count = 0
                    };
                }

                // Получаем все рабочие наборы
                var allWorksets = GetAllWorksets(doc);

                var result = new List<object>();
                int processedCount = 0;

                foreach (var workset in allWorksets)
                {
                    try
                    {
                        // Получаем переопределение видимости для рабочего набора на данном виде
                        WorksetVisibility visibility = view.GetWorksetVisibility(workset.Id);

                        // Определяем итоговый статус видимости
                        string finalVisibilityStatus;
                        bool isActuallyVisible;

                        switch (visibility)
                        {
                            case WorksetVisibility.Visible:
                                finalVisibilityStatus = "показать";
                                isActuallyVisible = true;
                                break;

                            case WorksetVisibility.Hidden:
                                finalVisibilityStatus = "скрыть";
                                isActuallyVisible = false;
                                break;

                            case WorksetVisibility.UseGlobalSetting:
                                // Для UseGlobalSetting нужно проверить глобальную настройку рабочего набора
                                bool isGloballyVisible = workset.IsVisibleByDefault;
                                finalVisibilityStatus = isGloballyVisible
                                    ? "использовать глобальную настройку видимости (видимый)"
                                    : "использовать глобальную настройку видимости (невидимый)";
                                isActuallyVisible = isGloballyVisible;
                                break;

                            default:
                                finalVisibilityStatus = visibility.ToString();
                                isActuallyVisible = false;
                                break;
                        }

                        // Формируем информацию о рабочем наборе
                        var worksetInfo = new
                        {
                            workset_id = workset.Id.IntegerValue,
                            workset_name = workset.Name ?? "Unnamed",
                            visibility_status = visibility.ToString(),
                            visibility_status_ru = finalVisibilityStatus,
                            globally_visible = workset.IsVisibleByDefault
                        };

                        result.Add(worksetInfo);
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result.Add(new
                        {
                            workset_id = workset.Id.IntegerValue,
                            workset_name = workset.Name ?? "Unnamed",
                            error = $"Ошибка: {ex.Message}"
                        });
                    }
                }

                return new
                {
                    view_id = viewId,
                    view_name = view.Name,
                    view_type = view.ViewType.ToString(),
                    workset_visibility = result,
                    count = result.Count
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении информации о видимости рабочих наборов: {ex.Message}",
                    workset_visibility = new List<object>(),
                    count = 0
                };
            }
        }

        /// <summary>
        /// Возвращает все рабочие наборы в документе
        /// </summary>
        private static List<Workset> GetAllWorksets(Document doc)
        {
            var worksets = new List<Workset>();

            try
            {
                // Получаем все рабочие наборы через FilteredWorksetCollector
                FilteredWorksetCollector collector = new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset);
                worksets = collector.ToWorksets().ToList();
                return worksets.OrderBy(w => w.Name).ToList();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Ошибка при получении рабочих наборов", ex);
            }

        }


        #endregion

        #region 33.3_get_link_graphics_overrides_in_view


        /// <summary>
        /// Возвращает информацию о графических переопределениях связанных файлов на указанном виде
        /// </summary>
        public static object GetLinkGraphicsOverridesInView(Document doc, int viewId)
        {
            // Получаем версию Revit
            int revitVersion = GetRevitVersion(doc);

            // Проверяем, поддерживается ли команда в текущей версии
            if (revitVersion < 2024)
            {
                return new
                {
                    error = $"Для данной версии Revit ({revitVersion}) не могу определить переопределение связей. " +
                            "Данная информация доступна только с Revit 2024.",
                    link_overrides = new List<object>(),
                    count = 0,
                    revit_version = revitVersion,
                    min_supported_version = 2024
                };
            }
            else
            {
#if R2020 || R2023
                return new
                {
                    error = $"Не удалось получить версию ревит. На данный момент получена версия: {revitVersion}",
                    link_overrides = new List<object>(),
                    count = 0,
                    revit_version = revitVersion,
                    min_supported_version = 2024
                };
#endif
            }



#if R2024
            try
            {

                // Далее идёт основной код для Revit 2024 и выше...
                // Получаем вид по ID
                ElementId viewElemId = IDHelper.ToElementId(viewId);
                View view = doc.GetElement(viewElemId) as View;

                if (view == null)
                {
                    return new
                    {
                        error = $"Вид с ID {viewId} не найден",
                        link_overrides = new List<object>(),
                        count = 0
                    };
                }

                // Проверяем, поддерживает ли вид переопределения связанных файлов
                if (!SupportsLinkOverrides(view))
                {
                    return new
                    {
                        error = $"Вид типа '{view.ViewType}' не поддерживает переопределения связанных файлов.",
                        link_overrides = new List<object>(),
                        count = 0
                    };
                }

                

                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                var result = new List<object>();

                foreach (var linkInstance in linkInstances)
                {
                    try
                    {
                        RevitLinkType linkType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                        string linkName = linkType?.Name ?? linkInstance.Name ?? "Unnamed";

                       
                        RevitLinkGraphicsSettings graphicsSettings = view.GetLinkOverrides(linkInstance.Id);

                        bool isHidden = linkInstance.IsHidden(view);

                        // Полутон
                        bool isHalftone = false;
                        try
                        {
                            OverrideGraphicSettings overrides = view.GetElementOverrides(linkInstance.Id);
                            isHalftone = overrides.Halftone;
                        }
                        catch { }

                        // Тип отображения
                        LinkVisibility visibilityType = graphicsSettings?.LinkVisibilityType ?? LinkVisibility.ByHostView;
                        ElementId linkedViewId = graphicsSettings?.LinkedViewId ?? ElementId.InvalidElementId;

                        string displaySetting;
                        object linkedViewInfo = null;

                        switch (visibilityType)
                        {
                            case LinkVisibility.ByHostView:
                                displaySetting = "по основному виду";
                                break;
                            case LinkVisibility.ByLinkView:
                                displaySetting = "по связанному виду";
                                if (linkedViewId != null && linkedViewId != ElementId.InvalidElementId)
                                {
                                    View linkedView = doc.GetElement(linkedViewId) as View;
                                    if (linkedView != null)
                                    {
                                        linkedViewInfo = new
                                        {
                                            view_id = linkedView.Id.IntegerValue,
                                            view_name = linkedView.Name ?? "Unnamed",
                                            view_type = linkedView.ViewType.ToString()
                                        };
                                    }
                                }
                                break;
                            case LinkVisibility.Custom:
                                displaySetting = "пользовательский";
                                break;
                            default:
                                displaySetting = visibilityType.ToString();
                                break;
                        }


                        result.Add(new
                        {
                            link_instance_id = linkInstance.Id.IntegerValue,
                            link_type_id = linkType?.Id.IntegerValue,
                            link_name = linkName,
                            is_hidden = isHidden,
                            is_halftone = isHalftone,
                            display_setting = displaySetting,
                            linked_view = linkedViewInfo
                        });
                    }
                    catch (Exception ex)
                    {
                        result.Add(new
                        {
                            link_instance_id = linkInstance.Id.IntegerValue,
                            link_name = linkInstance.Name ?? "Unnamed",
                            error = $"Ошибка: {ex.Message}"
                        });
                    }
                }

                return new
                {
                    success = true,
                    revit_version = revitVersion,
                    view_id = viewId,
                    view_name = view.Name,
                    view_type = view.ViewType.ToString(),
                    link_overrides = result,
                    count = result.Count
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении переопределений связанных файлов: {ex.Message}",
                    link_overrides = new List<object>(),
                    count = 0
                };
            }
#endif

            return new
            {
                error = "Link graphics overrides are available only in Revit 2024 configuration.",
                link_overrides = new List<object>(),
                count = 0,
                revit_version = revitVersion,
                min_supported_version = 2024
            };

        }

        /// <summary>
        /// Проверяет, поддерживает ли вид переопределения связанных файлов
        /// </summary>
        private static bool SupportsLinkOverrides(View view)
        {
            // Виды, которые поддерживают переопределения связанных файлов[citation:1][citation:4]
            var supportedTypes = new HashSet<ViewType>
            {
                ViewType.FloorPlan,          // Планы этажей
                ViewType.CeilingPlan,        // Планы потолков
                ViewType.ThreeD,             // 3D-виды
                ViewType.DraftingView,       // Чертёжные виды
                ViewType.Detail,             // Детальные виды
                ViewType.AreaPlan,           // Планы зон
                ViewType.Section,            // Разрезы
                ViewType.Elevation           // Фасады
            };

            return supportedTypes.Contains(view.ViewType);
        }

        /// <summary>
        /// Возвращает версию Revit в формате (год выпуска)
        /// </summary>
        private static int GetRevitVersion(Document doc)
        {
            try
            {
                // Получаем версию через Application
                string versionName = doc.Application.VersionName;

                // Примеры: "2024", "2023", "2022" и т.д.
                if (int.TryParse(versionName, out int version))
                {
                    return version;
                }

                // Альтернативный способ: через VersionNumber
                string versionNumber = doc.Application.VersionNumber;
                if (int.TryParse(versionNumber, out int versionNum))
                {
                    return versionNum;
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }


        #endregion

        #region 34_get_viewports_and_schedules_on_sheets   

        public static object GetViewportsAndSchedulesOnSheets(Document doc, List<int> sheetElementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (sheetElementIds == null || sheetElementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список sheetElementIds пуст или не указан",
                        sheet_contents = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int notSheetCount = 0;

                foreach (int sheetId in sheetElementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(sheetId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            result[sheetId] = new
                            {
                                error = "Лист не найден в документе"
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, является ли элемент листом (ViewSheet)
                        ViewSheet sheet = element as ViewSheet;
                        if (sheet == null)
                        {
                            result[sheetId] = new
                            {
                                error = "Элемент не является листом. Укажите ID листа (ViewSheet)."
                            };
                            notSheetCount++;
                            continue;
                        }

                        var sheetContents = new List<object>();

                        // 1. Получаем все видовые экраны (Viewport) на листе 
                        ICollection<ElementId> viewportIds = sheet.GetAllViewports();

                        foreach (ElementId viewportId in viewportIds)
                        {
                            Viewport viewport = doc.GetElement(viewportId) as Viewport;
                            if (viewport == null) continue;

                            ElementId referencedViewId = viewport.ViewId;
                            View referencedView = doc.GetElement(referencedViewId) as View;




                            sheetContents.Add(new
                            {
                                viewportId = IDHelper.ElIdInt(viewportId),
                                referencedViewId = IDHelper.ElIdInt(referencedViewId),
                                viewName = referencedView?.Name ?? "Unknown",
                                type = DetermineViewportType(referencedView)
                            });
                        }

                        // 2. Получаем все спецификации и легенды на листе
                        // Используем фильтрованный сборщик для поиска ScheduleSheetInstance на листе
                        // ScheduleSheetInstance — это элемент, представляющий спецификацию, размещённую на листе
                        var scheduleInstances = new FilteredElementCollector(doc, sheet.Id)
                            .OfClass(typeof(ScheduleSheetInstance))
                            .Cast<ScheduleSheetInstance>()
                            .ToList();

                        foreach (ScheduleSheetInstance scheduleInstance in scheduleInstances)
                        {
                            // Получаем ссылочный вид (саму спецификацию)
                            ElementId scheduleViewId = scheduleInstance.ScheduleId;
                            ViewSchedule schedule = doc.GetElement(scheduleViewId) as ViewSchedule;

                            // Определяем тип (спецификация или легенда)
                            string type = "Спецификация";

                            sheetContents.Add(new
                            {
                                viewportId = IDHelper.ElIdInt(scheduleInstance.Id),
                                referencedViewId = IDHelper.ElIdInt(scheduleViewId),
                                viewName = schedule?.Name ?? "Unknown",
                                type = type
                            });
                        }

                        // 3. Получаем оставшиеся типы
                        var allElementsOnSheet = new FilteredElementCollector(doc, sheet.Id)
                            .WhereElementIsNotElementType()
                            .ToElements();

                        // Создаём HashSet ID уже добавленных элементов (видовые экраны + спецификации)
                        var addedElementIds = new HashSet<int>();

                        // Добавляем ID видовых экранов
                        foreach (var viewportId in viewportIds)
                            addedElementIds.Add(IDHelper.ElIdInt(viewportId));

                        // Добавляем ID спецификаций
                        foreach (var scheduleInstance in scheduleInstances)
                            addedElementIds.Add(IDHelper.ElIdInt(scheduleInstance.Id));

                        // Проходим по всем элементам на листе и добавляем те, которые ещё не добавлены
                        foreach (Element elem in allElementsOnSheet)
                        {
                            int ElemId = IDHelper.ElIdInt(elem.Id);

                            // Пропускаем уже добавленные элементы
                            if (addedElementIds.Contains(ElemId))
                                continue;

                            // Определяем тип элемента для вывода
                            string elementType = GetElementTypeDescription(elem);

                            sheetContents.Add(new
                            {
                                viewportId = ElemId,                     // ID элемента
                                referencedViewId = (int?)null,           // У таких элементов нет ссылочного вида
                                viewName = GetElementDisplayName(elem),  // Имя элемента (если есть)
                                type = elementType
                            });
                        }

                        // Сортируем содержимое по типу и имени для удобства
                        var sortedContents = sheetContents
                            .OrderBy(c => c.GetType().GetProperty("type")?.GetValue(c, null))
                            .ThenBy(c => c.GetType().GetProperty("viewName")?.GetValue(c, null))
                            .ToList();

                        result[sheetId] = new
                        {
                            sheet_name = sheet.Name,
                            sheet_number = sheet.SheetNumber,
                            contents = sortedContents,
                            count = sortedContents.Count,
                            viewports_count = viewportIds.Count,
                            schedules_count = scheduleInstances.Count
                        };

                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[sheetId] = new
                        {
                            error = $"Ошибка при получении содержимого листа: {ex.Message}"
                        };
                    }
                }

                return new
                {
                    sheet_contents = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    not_sheets = notSheetCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении содержимого листов: {ex.Message}",
                    sheet_contents = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }


        private static string DetermineViewportType(View view)
        {
            switch (view.ViewType)
            {
                // Планы
                case ViewType.FloorPlan:
                    return "План этажа";  // План этажа
                case ViewType.CeilingPlan:
                    return "План потолка";  // План потолка
                case ViewType.EngineeringPlan:
                    return "План несущих конструкций";  // План несущих конструкций (Structural Plan)[citation:1][citation:5]
                case ViewType.AreaPlan:
                    return "План зонирования";  // План зон

                // Разрезы и фасады
                case ViewType.Section:
                    return "Разрез";  // Разрез[citation:3][citation:5]
                case ViewType.Elevation:
                    return "Фасад";  // Фасад[citation:3][citation:5]

                // 3D виды
                case ViewType.ThreeD:
                    return "3D вид";  // 3D вид
                case ViewType.Walkthrough:
                    return "Обход";  // Обход (3D-анимация)
                case ViewType.Rendering:
                    return "Камера";  // Визуализация

                // Легенды и чертежные виды
                case ViewType.Legend:
                    return "Легенда";  // Легенда[citation:3]
                case ViewType.DraftingView:
                    return "Чертёжный вид";  // Чертёжный вид[citation:3][citation:5]
                case ViewType.Detail:
                    return "Фрагмент";  // Фрагмент

                // Спецификации и листы
                case ViewType.Schedule:
                    return "Спецификация";  // Спецификация[citation:3][citation:5]
                case ViewType.DrawingSheet:
                    return "Лист";  // Лист
                case ViewType.ColumnSchedule:
                    return "Графическая спецификация колонн";  // Спецификация колонн
                case ViewType.PanelSchedule:
                    return "Спецификация панелей";  // Спецификация панелей

                //// Отчёты
                //case ViewType.CostReport:
                //    return "CostReport";  // Отчёт по стоимости
                //case ViewType.LoadsReport:
                //    return "LoadsReport";  // Отчёт по нагрузкам
                //case ViewType.PresureLossReport:
                //    return "PressureLossReport";  // Отчёт по потерям давления

                default:
                    return view.ViewType.ToString();  // Для остальных типов
            }
        }


        /// <summary>
        /// Возвращает понятное описание типа элемента на русском языке
        /// </summary>
        private static string GetElementTypeDescription(Element elem)
        {
            if (elem == null) return "Неизвестный элемент";

            // Текстовые заметки
            if (elem is TextNote) return "Текст";
            if (elem is TextElement) return "Текст";

            // Размеры
            if (elem is Dimension) return "Размер";

            // Линии и детализация
            if (elem is DetailLine) return "Линия детализации";
            if (elem is DetailArc) return "Дуга детализации";
            if (elem is DetailCurve) return "Кривая детализации";

            // Облака ревизий
            if (elem is RevisionCloud) return "Облако ревизии";

            // Штампы
            if (elem.Category != null && IDHelper.ElIdInt(elem.Category.Id) == (int)BuiltInCategory.OST_TitleBlocks) return "Основная надпись (штамп)";

            // Изображения
            if (elem is ImageInstance) return "Изображение";

            // Теги и выноски
            if (elem is IndependentTag) return "Текст с выноской";

            // Загружаемые семейства
            if (elem is FamilyInstance)
            {
                // Можно получить имя семейства для более детальной информации
                string familyName = (elem as FamilyInstance)?.Symbol?.Family?.Name;
                return string.IsNullOrEmpty(familyName) ? "Семейство" : $"Семейство: {familyName}";
            }

            // Символы штампов (могут быть на листе)
            if (elem is FamilySymbol) return "Символ семейства";

            // Проверка по категории
            if (elem.Category != null)
            {
                string catName = elem.Category.Name;
                return catName;
            }

            // Fallback
            return elem.GetType().Name;
        }

        /// <summary>
        /// Получает отображаемое имя элемента (для разных типов)
        /// </summary>
        private static string GetElementDisplayName(Element elem)
        {
            if (elem == null) return "Unknown";

            // У текстовых заметок нет имени, используем текст
            if (elem is TextNote textNote)
            {
                string text = textNote.Text;
                return string.IsNullOrEmpty(text) ? "Текст" : text;
            }

            // У размеров может быть значение
            if (elem is Dimension dim)
            {
                string value = dim.ValueString;
                return $"Размер: {value}";
            }

            // У семейств используем имя типа
            if (elem is FamilyInstance fi)
            {
                return fi.Symbol?.Name ?? fi.Name ?? "Семейство";
            }

            // У штампов получаем номер листа через параметр
            if (elem.Category != null && IDHelper.ElIdInt(elem.Category.Id) == (int)BuiltInCategory.OST_TitleBlocks)
            {
                Parameter sheetNumParam = elem.get_Parameter(BuiltInParameter.SHEET_NUMBER);
                if (sheetNumParam != null && sheetNumParam.HasValue)
                {
                    return $"Штамп: {sheetNumParam.AsString()}";
                }
                return "Штамп";
            }

            // Стандартное имя элемента
            return !string.IsNullOrEmpty(elem.Name) ? elem.Name : elem.GetType().Name;
        }



        #endregion

        #region 35_get_schedules_info_and_columns

        public static object GetSchedulesInfoAndColumns(Document doc, List<int> scheduleElementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (scheduleElementIds == null || scheduleElementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список scheduleElementIds пуст или не указан",
                        schedules_info = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int notScheduleCount = 0;

                foreach (int scheduleId in scheduleElementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(scheduleId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            result[scheduleId] = new
                            {
                                error = "Спецификация не найдена в документе"
                            };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, является ли элемент спецификацией (ViewSchedule)
                        ViewSchedule schedule = element as ViewSchedule;
                        if (schedule == null)
                        {
                            result[scheduleId] = new
                            {
                                error = "Элемент не является спецификацией. Укажите ID спецификации (ViewSchedule)."
                            };
                            notScheduleCount++;
                            continue;
                        }

                        // Получаем определение спецификации
                        ScheduleDefinition definition = schedule.Definition;
                        if (definition == null)
                        {
                            result[scheduleId] = new
                            {
                                error = "Не удалось получить определение спецификации",
                                scheduleName = schedule.Name
                            };
                            continue;
                        }

                        // Получаем количество строк
                        int rowCount = 0;
                        try
                        {
                            TableData tableData = schedule.GetTableData();
                            if (tableData != null)
                            {
                                TableSectionData bodySection = tableData.GetSectionData(SectionType.Body);
                                if (bodySection != null)
                                {
                                    rowCount = bodySection.NumberOfRows;
                                }
                            }
                        }
                        catch
                        {
                            rowCount = -1; // Если не удалось получить количество строк
                        }





                        // Получаем все поля (SchedulableField) из определения
                        // ScheduleDefinition содержит коллекцию полей, которые становятся столбцами спецификации
                        IList<ScheduleFieldId> fieldIds = definition.GetFieldOrder();
                        var columns = new List<object>();
                        var filtersList = new List<object>();

                        // Сначала создадим словарь полей для быстрого доступа (понадобится для фильтров)
                        var fieldsMap = new Dictionary<int, ScheduleField>();

                        foreach (ScheduleFieldId fieldId in fieldIds)
                        {
                            // Получаем ScheduleField (реальный столбец спецификации)
                            ScheduleField field = definition.GetField(fieldId);
                            fieldsMap[fieldId.IntegerValue] = field;

                            // Заголовок столбца (то, что видит пользователь)
                            string header = field.ColumnHeading;

                            // Имя параметра (через GetName())
                            string parameterName = field.GetName();

                            // ID параметра (только для обычных полей, не для вычисляемых)
                            int parameterId = -1;

                            //тип параметра (фомрула, процент и т.п.)
                            string fieldTypeDescription = GetFieldTypeDescription(field.FieldType);

                            // Информация о вычисляемых полях
                            string calculatedType = null;
                            string percentageOfField = null;
                            string percentageByField = null;
                            IList<object> combinedParameters = null;

                            if (field.IsCalculatedField)
                            {
                                // Определяем тип вычисляемого поля
                                switch (field.FieldType)
                                {
                                    case ScheduleFieldType.Formula:
                                        calculatedType = "Formula";
                                        break;
                                    case ScheduleFieldType.Percentage:
                                        calculatedType = "Percentage";
                                        // Для Percentage полей получаем информацию о том, от какого поля считается процент
                                        if (field.PercentageOf != null && field.PercentageOf.IntegerValue != -1)
                                        {
                                            if (fieldsMap.ContainsKey(field.PercentageOf.IntegerValue))
                                            {
                                                ScheduleField sourceField = fieldsMap[field.PercentageOf.IntegerValue];
                                                percentageOfField = sourceField?.GetName() ?? "Unknown";
                                            }
                                            else
                                            {
                                                percentageOfField = field.PercentageOf.IntegerValue.ToString();
                                            }
                                        }
                                        // Получаем информацию о группировке для процентного поля
                                        if (field.PercentageBy != null && field.PercentageBy.IntegerValue != -1)
                                        {
                                            if (fieldsMap.ContainsKey(field.PercentageBy.IntegerValue))
                                            {
                                                ScheduleField groupField = fieldsMap[field.PercentageBy.IntegerValue];
                                                percentageByField = groupField?.GetName() ?? "Unknown";
                                            }
                                            else
                                            {
                                                percentageByField = field.PercentageBy.IntegerValue.ToString();
                                            }
                                        }
                                        break;
                                    case ScheduleFieldType.Count:
                                        calculatedType = "Count";
                                        break;
                                    default:
                                        calculatedType = field.FieldType.ToString();
                                        break;
                                }
                            }
                            // Для CombinedParameter полей получаем список участвующих параметров
                            if (field.FieldType == ScheduleFieldType.CombinedParameter && field.IsCombinedParameterField)
                            {
                                try
                                {
                                    // ✅ Правильный тип: IList<TableCellCombinedParameterData>
                                    IList<TableCellCombinedParameterData> combinedParamData = field.GetCombinedParameters();
                                    if (combinedParamData != null && combinedParamData.Count > 0)
                                    {
                                        combinedParameters = new List<object>();
                                        foreach (TableCellCombinedParameterData paramData in combinedParamData)
                                        {
                                            // Получаем ID параметра
                                            ElementId paramId = paramData.ParamId;

                                            // Получаем имя параметра
                                            string paramName = null;
                                            Element paramElement = doc.GetElement(paramId);
                                            if (paramElement != null && paramElement is ParameterElement)
                                            {
                                                paramName = paramElement.Name;
                                            }

                                            // Пробуем получить префикс, суффикс и образец
                                            string prefix = paramData.Prefix;
                                            string separator = paramData.Separator;
                                            string suffix = paramData.Suffix;
                                            string sample = paramData.SampleValue;

                                            combinedParameters.Add(new
                                            {
                                                parameterId = IDHelper.ElIdInt(paramId),
                                                parameterName = paramName ?? $"Unknown_{IDHelper.ElIdInt(paramId)}",
                                                prefix = prefix,
                                                separator = separator,
                                                suffix = suffix,
                                                sample = sample
                                            });
                                        }
                                    }
                                }
                                catch { }
                            }

                            // Для обычных полей получаем ParameterId
                            if (!field.IsCalculatedField && !field.IsCombinedParameterField && field.HasSchedulableField)
                            {
                                try
                                {
                                    SchedulableField schedulableField = field.GetSchedulableField();
                                    if (schedulableField != null && schedulableField.ParameterId != null)
                                    {
                                        parameterId = IDHelper.ElIdInt(schedulableField.ParameterId);
                                    }
                                }
                                catch { }
                            }





                            columns.Add(new
                            {
                                header = header ?? parameterName,
                                parameterName = parameterName,
                                parameterId = parameterId,
                                isHidden = field.IsHidden,           // ✅ Теперь IsHidden доступен
                                isCalculated = field.IsCalculatedField,
                                isCombinedParameter = field.IsCombinedParameterField,
                                fieldType = field.FieldType.ToString(),
                                fieldTypeDescription = fieldTypeDescription,
                                calculatedType = calculatedType,
                                percentageOfField = percentageOfField,
                                percentageByField = percentageByField,
                                combinedParameters = combinedParameters
                            });
                        }

                        //Получаем фильтры спецификации
                        IList<ScheduleFilter> filters = definition.GetFilters();

                        foreach (ScheduleFilter filter in filters)
                        {
                            // Получаем информацию о поле, к которому применён фильтр
                            ScheduleField filterField = null;
                            string fieldName = null;
                            int fieldParameterId = -1;

                            if (fieldsMap.ContainsKey(filter.FieldId.IntegerValue))
                            {
                                filterField = fieldsMap[filter.FieldId.IntegerValue];
                                fieldName = filterField.GetName();

                                if (!filterField.IsCalculatedField && filterField.HasSchedulableField)
                                {
                                    try
                                    {
                                        SchedulableField sf = filterField.GetSchedulableField();
                                        if (sf != null && sf.ParameterId != null)
                                        {
                                            fieldParameterId = IDHelper.ElIdInt(sf.ParameterId);
                                        }
                                    }
                                    catch { }
                                }
                            }

                            // Получаем значение фильтра в зависимости от его типа
                            object filterValue = null;
                            string filterValueType = null;

                            if (filter.IsStringValue)
                            {
                                filterValue = filter.GetStringValue();
                                filterValueType = "String";
                            }
                            else if (filter.IsIntegerValue)
                            {
                                filterValue = filter.GetIntegerValue();
                                filterValueType = "Integer";
                            }
                            else if (filter.IsDoubleValue)
                            {
                                filterValue = filter.GetDoubleValue();
                                filterValueType = "Double";
                            }
                            else if (filter.IsElementIdValue)
                            {
                                int elemIdValue = IDHelper.ElIdInt(filter.GetElementIdValue());
                                filterValue = elemIdValue;
                                filterValueType = "ElementId";

                                // Для ElementId можно попробовать получить имя элемента
                                Element referencedElem = doc.GetElement(IDHelper.ToElementId(elemIdValue));
                                if (referencedElem != null)
                                {
                                    filterValue = new { id = elemIdValue, name = referencedElem.Name ?? referencedElem.GetType().Name };
                                }
                            }
                            else if (filter.IsNullValue)
                            {
                                filterValue = null;
                                filterValueType = "Null (HasParameter)";
                            }

                            filtersList.Add(new
                            {
                                fieldId = filter.FieldId.IntegerValue,
                                fieldName = fieldName ?? "Unknown",
                                fieldParameterId = fieldParameterId,
                                filterType = filter.FilterType.ToString(),  // [citation:10]
                                filterTypeDescription = GetFilterTypeDescription(filter.FilterType),
                                value = filterValue,
                                valueType = filterValueType
                            });
                        }





                        // Дополнительная информация о спецификации
                        var scheduleInfo = new
                        {
                            scheduleId = IDHelper.ElIdInt(schedule.Id),
                            scheduleName = schedule.Name ?? "Unnamed",
                            categoryId = definition.CategoryId != null ? IDHelper.ElIdInt(definition.CategoryId) : -1,
                            categoryName = GetCategoryName(doc, definition.CategoryId),
                            columns = columns,
                            columnsCount = columns.Count,
                            rowCount = rowCount,
                            filters = filtersList,
                            filtersCount = filtersList.Count,
                            hasFilters = filtersList.Count > 0
                        };

                        result[scheduleId] = scheduleInfo;
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[scheduleId] = new
                        {
                            error = $"Ошибка при получении структуры спецификации: {ex.Message}",
                            scheduleId = scheduleId
                        };
                    }
                }

                return new
                {
                    schedules_info = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    not_schedules = notScheduleCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении информации о спецификациях: {ex.Message}",
                    schedules_info = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        /// <summary>
        /// Возвращает понятное описание типа поля спецификации
        /// </summary>
        private static string GetFieldTypeDescription(ScheduleFieldType fieldType)
        {
            switch (fieldType)
            {
                case ScheduleFieldType.Instance:
                    return "Параметр экземпляра";
                case ScheduleFieldType.ElementType:
                    return "Параметр типа";
                case ScheduleFieldType.Count:
                    return "Количество элементов в строке";
                case ScheduleFieldType.Formula:
                    return "Формула";
                case ScheduleFieldType.Percentage:
                    return "Процентное отношение";
                case ScheduleFieldType.CombinedParameter:
                    return "Объединённый параметр";
                case ScheduleFieldType.ViewBased:
                    return "Параметр зависящий от вида (площадь/периметр помещения, номер ревизии и т.п.)";
                case ScheduleFieldType.Analytical:
                    return "Аналитический параметр";
                default:
                    return fieldType.ToString();
            }
        }






        /// <summary>
        /// Возвращает имя категории по номеру id
        /// </summary>
        private static string GetCategoryName(Document doc, ElementId categoryId)
        {
            if (categoryId == null || IDHelper.ElIdInt(categoryId) == -1)
                return null;

            try
            {
                Category category = Category.GetCategory(doc, categoryId);
                return category?.Name;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Возвращает понятное описание типа фильтра
        /// </summary>
        private static string GetFilterTypeDescription(ScheduleFilterType filterType)
        {
            switch (filterType)
            {
                case ScheduleFilterType.Equal: return "равно";
                case ScheduleFilterType.NotEqual: return "не равно";
                case ScheduleFilterType.GreaterThan: return "больше";
                case ScheduleFilterType.GreaterThanOrEqual: return "больше или равно";
                case ScheduleFilterType.LessThan: return "меньше";
                case ScheduleFilterType.LessThanOrEqual: return "меньше или равно";
                case ScheduleFilterType.Contains: return "содержит";
                case ScheduleFilterType.BeginsWith: return "начинается с";
                case ScheduleFilterType.EndsWith: return "заканчивается на";
                case ScheduleFilterType.NotContains: return "не содержит";
                case ScheduleFilterType.NotBeginsWith: return "не начинается с";
                case ScheduleFilterType.NotEndsWith: return "не заканчивается на";
                default: return filterType.ToString();
            }
        }


        #endregion

        #region 35.1_get_schedule_sorting_info

        /// <summary>
        /// Возвращает информацию о правилах сортировки в спецификации Revit
        /// </summary>
        public static object GetScheduleSortingInfo(Document doc, List<int> scheduleElementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (scheduleElementIds == null || scheduleElementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список scheduleElementIds пуст или не указан",
                        schedules_sorting = new Dictionary<int, object>(),
                        count = 0
                    };
                }

                var result = new Dictionary<int, object>();
                int processedCount = 0;
                int notFoundCount = 0;
                int notScheduleCount = 0;

                foreach (int scheduleId in scheduleElementIds.Distinct())
                {
                    try
                    {
                        ElementId elemId = IDHelper.ToElementId(scheduleId);
                        Element element = doc.GetElement(elemId);

                        if (element == null)
                        {
                            result[scheduleId] = new { error = "Спецификация не найдена в документе" };
                            notFoundCount++;
                            continue;
                        }

                        // Проверяем, является ли элемент спецификацией (ViewSchedule)
                        ViewSchedule schedule = element as ViewSchedule;
                        if (schedule == null)
                        {
                            result[scheduleId] = new { error = "Элемент не является спецификацией" };
                            notScheduleCount++;
                            continue;
                        }

                        ScheduleDefinition definition = schedule.Definition;
                        if (definition == null)
                        {
                            result[scheduleId] = new { error = "Не удалось получить определение спецификации" };
                            continue;
                        }

                        // ========== ПОЛУЧАЕМ ИНФОРМАЦИЮ О СОРТИРОВКЕ ==========

                        // 1. Получаем порядок сортировки (список полей, по которым идёт сортировка)
                        IList<ScheduleSortGroupField> sortGroups = definition.GetSortGroupFields();

                        var sortingInfo = new List<object>();

                        for (int i = 0; i < sortGroups.Count; i++)
                        {
                            ScheduleSortGroupField sortField = sortGroups[i];

                            // Получаем поле, по которому идёт сортировка
                            ScheduleField field = null;
                            string fieldName = null;
                            int fieldId = -1;
                            int parameterId = -1;
                            string fieldType = null;

                            try
                            {
                                fieldId = sortField.FieldId.IntegerValue;
                                field = definition.GetField(sortField.FieldId);
                                if (field != null)
                                {
                                    fieldName = field.GetName();
                                    fieldType = field.FieldType.ToString();

                                    // Получаем parameterId для обычных полей
                                    if (!field.IsCalculatedField && field.HasSchedulableField)
                                    {
                                        SchedulableField sf = field.GetSchedulableField();
                                        if (sf != null && sf.ParameterId != null)
                                        {
                                            parameterId = IDHelper.ElIdInt(sf.ParameterId);
                                        }
                                    }
                                }
                            }
                            catch { }

                            // Получаем направление сортировки
                            string sortOrder = GetSortOrderDescription(sortField.SortOrder);

                            sortingInfo.Add(new
                            {
                                level = i + 1,  // Уровень сортировки (1, 2, 3...)
                                fieldId = fieldId,
                                fieldName = fieldName ?? "Unknown",
                                parameterId = parameterId,
                                fieldType = fieldType,
                                //Правила сортировки по параметрам
                                sortOrder = sortOrder,
                                showBlankRow = sortField.ShowBlankLine,
                                showHeader = sortField.ShowHeader,
                                showFooter = sortField.ShowFooter,
                                showFooterCount = sortField.ShowFooterCount,
                                showFooterTitle = sortField.ShowFooterTitle,
                                //Общий итог
                                showGrandTotal = definition.ShowGrandTotal,
                                showGrandTotalCount = definition.ShowGrandTotalCount,
                                showGrandTotalTitle = definition.ShowGrandTotalTitle,
                                grandTitle = definition.GrandTotalTitle,
                                //Для каждого экземпляра или нет
                                isItemized = definition.IsItemized
                            });
                        }

                        // ========== ДОПОЛНИТЕЛЬНАЯ ИНФОРМАЦИЯ О СОРТИРОВКЕ ==========

                        // Проверяем, используется ли сортировка по умолчанию
                        bool hasCustomSorting = sortGroups.Count > 0;

                        var scheduleSortingInfo = new
                        {
                            scheduleId = IDHelper.ElIdInt(schedule.Id),
                            scheduleName = schedule.Name ?? "Unnamed",
                            hasSorting = hasCustomSorting,
                            sortLevelsCount = sortGroups.Count,
                            sorting = sortingInfo,
                            note = hasCustomSorting
                                ? "Спецификация имеет пользовательскую сортировку"
                                : "Спецификация использует сортировку по умолчанию"
                        };

                        result[scheduleId] = scheduleSortingInfo;
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        result[scheduleId] = new
                        {
                            error = $"Ошибка при получении информации о сортировке: {ex.Message}",
                            scheduleId = scheduleId
                        };
                    }
                }

                return new
                {
                    schedules_sorting = result,
                    count = result.Count,
                    processed_successfully = processedCount,
                    not_found = notFoundCount,
                    not_schedules = notScheduleCount
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при получении информации о сортировке: {ex.Message}",
                    schedules_sorting = new Dictionary<int, object>(),
                    count = 0
                };
            }
        }

        /// <summary>
        /// Возвращает понятное описание направления сортировки
        /// </summary>
        private static string GetSortOrderDescription(ScheduleSortOrder sortOrder)
        {
            switch (sortOrder)
            {
                case ScheduleSortOrder.Ascending:
                    return "По возрастанию (А → Я, 0 → 9)";
                case ScheduleSortOrder.Descending:
                    return "По убыванию (Я → А, 9 → 0)";
                default:
                    return sortOrder.ToString();
            }
        }

        #endregion

        //#region 35.2_get_schedule_rows_with_elements

        //public static object GetScheduleRowsWithElements(Document doc, int scheduleId)
        //{
        //    try
        //    {
        //        Element element = doc.GetElement(new ElementId(scheduleId));
        //        if (element == null)
        //        {
        //            return new { error = $"Спецификация с ID {scheduleId} не найдена", success = false };
        //        }

        //        ViewSchedule schedule = element as ViewSchedule;
        //        if (schedule == null)
        //        {
        //            return new { error = "Элемент не является спецификацией", success = false };
        //        }

        //        TableData tableData = schedule.GetTableData();
        //        if (tableData == null)
        //        {
        //            return new { error = "Не удалось получить данные таблицы", success = false };
        //        }

        //        TableSectionData bodySection = tableData.GetSectionData(SectionType.Body);
        //        if (bodySection == null)
        //        {
        //            return new { error = "Не удалось получить секцию Body", success = false };
        //        }


        //        ScheduleDefinition definition = schedule.Definition;
        //        if (definition == null)
        //        {
        //            return new { error = "Не удалось получить определение спецификации", success = false };
        //        }

        //        // ========== 1. ПОЛУЧАЕМ ИНДЕКСЫ ВИДИМЫХ СТОЛБЦОВ ==========
        //        IList<ScheduleFieldId> fieldOrder = definition.GetFieldOrder();
        //        var visibleColumnIndices = new List<int>();      // Реальные индексы видимых столбцов
        //        var visibleFieldIds = new List<int>();           // ID полей для информации

        //        for (int i = 0; i < fieldOrder.Count; i++)
        //        {
        //            ScheduleField field = definition.GetField(fieldOrder[i]);
        //            if (field != null && !field.IsHidden)
        //            {
        //                visibleColumnIndices.Add(i);              // Сохраняем РЕАЛЬНЫЙ индекс
        //                visibleFieldIds.Add(fieldOrder[i].IntegerValue);
        //            }
        //        }

        //        // ========== 2. ПОЛУЧАЕМ ЗАГОЛОВКИ ДЛЯ ВИДИМЫХ СТОЛБЦОВ ==========
        //        var headers = new List<string>();
        //        foreach (int realColIndex in visibleColumnIndices)
        //        {
        //            string header = schedule.GetCellText(SectionType.Header, 0, realColIndex);
        //            headers.Add(string.IsNullOrEmpty(header) ? $"Column_{realColIndex}" : header);
        //        }



        //        //=====================================
        //        int rowCount = bodySection.NumberOfRows;
        //        int visibleColumnCount = visibleColumnIndices.Count;


        //        // Получаем все элементы, связанные со спецификацией
        //        var allElementsInSchedule = new FilteredElementCollector(doc, schedule.Id)
        //            .WhereElementIsNotElementType()
        //            .ToElementIds()
        //            .ToList();


        //        //// Определяем, сгруппирована ли спецификация
        //        //bool isItemized = schedule.Definition?.IsItemized ?? true;

        //        var rows = new List<object>();
        //        for (int row = 0; row < rowCount; row++)
        //        {
        //            var rowValues = new Dictionary<string, string>();

        //            // Используем реальные индексы столбцов!
        //            for (int colIdx = 0; colIdx < visibleColumnCount; colIdx++)
        //            {
        //                int realColIndex = visibleColumnIndices[colIdx];

        //                CellType cellType = bodySection.GetCellType(row, realColIndex);
        //                switch (cellType)
        //                {
        //                    case CellType.Text:
        //                    case CellType.ParameterText:
        //                        // Работает GetCellText
        //                        string textValue = schedule.GetCellText(SectionType.Body, row, realColIndex);
        //                        break;

        //                    //case CellType.Parameter:
        //                    //    // Для числовых параметров используем GetCellValue
        //                    //    object numericValue = schedule.GetCellValue(SectionType.Body, row, realColIndex);
        //                    //    break;

        //                    case CellType.CalculatedValue:
        //                        // Для вычисляемых полей
        //                        string calcValue = schedule.GetCalculatedValueText(SectionType.Body, row, realColIndex);
        //                        break;
        //                }



        //                string cellValue = schedule.GetCellText(SectionType.Body, row, realColIndex);
        //                rowValues[headers[colIdx]] = cellValue ?? "";
        //            }

        //            rows.Add(new
        //            {
        //                row_index = row,
        //                values = rowValues
        //            });
        //        }


        //        return new
        //        {
        //            schedule_id = schedule.Id.IntegerValue,
        //            schedule_name = schedule.Name ?? "Unnamed",
        //            is_itemized = definition.IsItemized,
        //            row_count = rowCount,
        //            column_count = visibleColumnCount,
        //            total_columns = fieldOrder.Count,
        //            hidden_columns_count = fieldOrder.Count - visibleColumnCount,
        //            headers = headers,
        //            rows = rows,
        //            element_ids = allElementsInSchedule
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new { error = ex.Message, success = false };
        //    }
        //}

        //#endregion

        #region 36_get_if_elements_pass_filter   

        public static object GetIfElementsPassFilter(Document doc, int filterId, List<int> elementIds)
        {
            try
            {
                // Проверяем входные параметры
                if (filterId == -1 || filterId == 0)
                {
                    return new
                    {
                        error = "Не указан filterId или передан некорректный ID",
                        filter_results = new Dictionary<int, bool>(),
                        count = 0
                    };
                }

                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        error = "Список elementIds пуст или не указан",
                        filter_results = new Dictionary<int, bool>(),
                        count = 0
                    };
                }

                // Получаем фильтр по ID
                ElementId filterElemId = IDHelper.ToElementId(filterId);
                ParameterFilterElement filter = doc.GetElement(filterElemId) as ParameterFilterElement;

                if (filter == null)
                {
                    return new
                    {
                        error = $"Фильтр с ID {filterId} не найден в документе. Убедитесь, что это корректный ID фильтра видов.",
                        filter_results = new Dictionary<int, bool>(),
                        count = 0
                    };
                }

                // Получаем ElementFilter (правила фильтрации)
                ElementFilter elementFilter = filter.GetElementFilter();

                if (elementFilter == null)
                {
                    return new
                    {
                        error = "Не удалось получить правила фильтрации. Возможно, фильтр не содержит условий.",
                        filter_results = new Dictionary<int, bool>(),
                        count = 0
                    };
                }

                // Собираем элементы для проверки
                var elementsToCheck = new List<Element>();
                var validElementIds = new List<int>();
                var notFoundIds = new List<int>();

                foreach (int elementId in elementIds.Distinct())
                {
                    ElementId elemId = IDHelper.ToElementId(elementId);
                    Element element = doc.GetElement(elemId);

                    if (element == null)
                    {
                        notFoundIds.Add(elementId);
                    }
                    else
                    {
                        elementsToCheck.Add(element);
                        validElementIds.Add(elementId);
                    }
                }

                // Используем FilteredElementCollector с элементами для проверки
                // Создаём пасс-фильтр, который будет проверять элементы из нашего списка
                var results = new Dictionary<int, bool>();

                // Проходим по каждому элементу и проверяем его через ElementFilter
                foreach (Element element in elementsToCheck)
                {
                    bool passes = elementFilter.PassesFilter(element);
                    results[IDHelper.ElIdInt(element.Id)] = passes;
                }

                //// Добавляем не найденные элементы с false
                //foreach (int notFoundId in notFoundIds)
                //{
                //    results[notFoundId] = false;
                //}

                // Получаем информацию о фильтре для контекста
                string filterName = filter.Name ?? "Unnamed";
                ICollection<ElementId> categoryIds = filter.GetCategories();
                var categoryInfo = new List<string>();
                foreach (ElementId catId in categoryIds)
                {
                    Category category = Category.GetCategory(doc, catId);
                    if (category != null)
                    {
                        categoryInfo.Add(category.Name);
                    }
                }

                // Подсчитываем статистику
                int passedCount = results.Count(r => r.Value);
                int failedCount = results.Count(r => !r.Value);

                return new
                {
                    filter_results = results,
                    count = results.Count,
                    passed_count = passedCount,
                    failed_count = failedCount,
                    not_found_count = notFoundIds.Count,
                    filter_info = new
                    {
                        filter_id = filterId,
                        filter_name = filterName,
                        categories = categoryInfo,
                        has_rules = elementFilter != null
                    }
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    error = $"Ошибка при проверке элементов на соответствие фильтру: {ex.Message}",
                    filter_results = new Dictionary<int, bool>(),
                    count = 0
                };
            }
        }

        #endregion

        #region 37_set_view_section_box_to_elements

        /// <summary>
        /// Подрезает 3D вид по границам указанных элементов с отступом
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="uiDoc">UI документ (для доступа к активному виду)</param>
        /// <param name="elementIds">Список ID элементов</param>
        /// <param name="marginMM">Отступ в миллиметрах (по умолчанию 500)</param>
        /// <returns>Результат операции</returns>
        public static object SetViewSectionBoxToElements(Document doc, UIDocument uiDoc, List<int> elementIds, double marginMM = 500)
        {
            try
            {
                // 1. Проверка: есть ли активный вид?
                if (uiDoc == null || uiDoc.ActiveView == null)
                {
                    return new
                    {
                        success = false,
                        error = "Нет активного вида. Пожалуйста, откройте вид в Revit."
                    };
                }

                View activeView = uiDoc.ActiveView;

                // 2. Проверка: является ли вид 3D?
                if (activeView.ViewType != ViewType.ThreeD)
                {
                    return new
                    {
                        success = false,
                        error = $"Открыт вид типа '{activeView.ViewType}'. Пожалуйста, откройте 3D вид и повторите команду.",
                        current_view_type = activeView.ViewType.ToString(),
                        current_view_name = activeView.Name
                    };
                }

                // 3. Проверка: есть ли элементы для обработки?
                if (elementIds == null || elementIds.Count == 0)
                {
                    return new
                    {
                        success = false,
                        error = "Список elementIds пуст. Укажите ID элементов для подрезки вида."
                    };
                }

                // 4. Получаем элементы по ID
                var elements = new List<Element>();
                var invalidIds = new List<int>();

                foreach (int id in elementIds.Distinct())
                {
                    Element elem = doc.GetElement(IDHelper.ToElementId(id));
                    if (elem == null)
                    {
                        invalidIds.Add(id);
                    }
                    else
                    {
                        elements.Add(elem);
                    }
                }

                if (elements.Count == 0)
                {
                    return new
                    {
                        success = false,
                        error = "Ни один из указанных ID не соответствует существующему элементу.",
                        invalid_ids = invalidIds
                    };
                }

                // 5. Вычисляем bounding box всех элементов
                BoundingBoxXYZ combinedBoundingBox = null;

                foreach (Element elem in elements)
                {
                    BoundingBoxXYZ elemBoundingBox = elem.get_BoundingBox(activeView);

                    if (elemBoundingBox == null)
                    {
                        continue; // У некоторых элементов может не быть геометрии
                    }

                    if (combinedBoundingBox == null)
                    {
                        // Создаём копию первого bounding box'а
                        combinedBoundingBox = new BoundingBoxXYZ();
                        combinedBoundingBox.Min = elemBoundingBox.Min;
                        combinedBoundingBox.Max = elemBoundingBox.Max;
                    }
                    else
                    {
                        // Расширяем границы
                        XYZ newMin = new XYZ(
                            Math.Min(combinedBoundingBox.Min.X, elemBoundingBox.Min.X),
                            Math.Min(combinedBoundingBox.Min.Y, elemBoundingBox.Min.Y),
                            Math.Min(combinedBoundingBox.Min.Z, elemBoundingBox.Min.Z)
                        );
                        XYZ newMax = new XYZ(
                            Math.Max(combinedBoundingBox.Max.X, elemBoundingBox.Max.X),
                            Math.Max(combinedBoundingBox.Max.Y, elemBoundingBox.Max.Y),
                            Math.Max(combinedBoundingBox.Max.Z, elemBoundingBox.Max.Z)
                        );
                        combinedBoundingBox.Min = newMin;
                        combinedBoundingBox.Max = newMax;
                    }
                }

                if (combinedBoundingBox == null)
                {
                    return new
                    {
                        success = false,
                        error = "Не удалось вычислить границы элементов. Возможно, элементы не имеют 3D-геометрии."
                    };
                }

                // 6. Переводим отступ из миллиметров в футы (Revit работает в футах)
                double marginFeet = marginMM / 304.8;

                // 7. Расширяем bounding box на отступ
                XYZ expandedMin = new XYZ(
                    combinedBoundingBox.Min.X - marginFeet,
                    combinedBoundingBox.Min.Y - marginFeet,
                    combinedBoundingBox.Min.Z - marginFeet
                );
                XYZ expandedMax = new XYZ(
                    combinedBoundingBox.Max.X + marginFeet,
                    combinedBoundingBox.Max.Y + marginFeet,
                    combinedBoundingBox.Max.Z + marginFeet
                );

                // 8. Создаём новый bounding box для section box
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Min = expandedMin;
                sectionBox.Max = expandedMax;

                // 9. Устанавливаем section box для вида
                // Включаем секционный параллелепипед
                View3D active3DView = activeView as View3D;

                using (Transaction t = new Transaction(doc, "Подрезать 3D вид по элементам"))
                {
                    t.Start();

                    active3DView.IsSectionBoxActive = true;
                    active3DView.SetSectionBox(sectionBox);

                    doc.Regenerate();

                    t.Commit();
                }

                uiDoc.RefreshActiveView();


                return new
                {
                    success = true,
                    message = $"Вид успешно подрезан по границам {elements.Count} элементов с отступом {marginMM} мм.",
                    elements_processed = elements.Count,
                    invalid_ids = invalidIds.Count > 0 ? invalidIds : null,
                    bounding_box = new
                    {
                        min = new { x = expandedMin.X, y = expandedMin.Y, z = expandedMin.Z },
                        max = new { x = expandedMax.X, y = expandedMax.Y, z = expandedMax.Z }
                    }
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    success = false,
                    error = $"Ошибка при подрезке вида: {ex.Message}"
                };
            }
        }

        #endregion

        #region 38_get_journal_entries_since

        /// <summary>
        /// Возвращает записи из журнала Revit начиная с указанной даты
        /// </summary>
        /// <param name="doc">Документ Revit (для получения версии)</param>
        /// <param name="userDateTime">Дата и время в формате (день.месяц.год час:минута:секунда)</param>
        public static object GetJournalEntriesSince(Document doc, string userDateTime)
        {
            try
            {
                // 1. Парсим введённую дату/время
                DateTime targetDateTime = ParseUserDateTime(userDateTime);

                if (targetDateTime == DateTime.MinValue)
                {
                    return new
                    {
                        success = false,
                        error = "Не удалось распознать дату. Используйте форматы: '26 марта', '26.03.2026', '26.03.2026 14:30'"
                    };
                }

                // 2. Определяем версию Revit
                string revitVersion = "Autodesk Revit " + GetRevitVersion(doc).ToString();
                if (string.IsNullOrEmpty(revitVersion))
                {
                    return new
                    {
                        success = false,
                        error = "Не удалось определить версию Revit"
                    };
                }

                // 3. Формируем путь к папке журналов
                string journalFolder = GetJournalFolderPath(revitVersion);
                if (!Directory.Exists(journalFolder))
                {
                    return new
                    {
                        success = false,
                        error = $"Папка с журналами не найдена: {journalFolder}"
                    };
                }

                // 4. Находим все файлы журналов
                var journalFiles = Directory.GetFiles(journalFolder, "journal.*.txt")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.LastWriteTime)
                    .ToList();

                if (journalFiles.Count == 0)
                {
                    return new
                    {
                        success = false,
                        error = "Файлы журналов не найдены"
                    };
                }

                // 5. Находим самый свежий журнал
                FileInfo latestJournal = journalFiles.First();

                // 6. СОЗДАЁМ КОПИЮ ФАЙЛА (чтобы обойти блокировку Revit)
                string tempCopyPath = Path.Combine(Path.GetTempPath(), $"journal_copy_{Guid.NewGuid()}.txt");

                try
                {
                    File.Copy(latestJournal.FullName, tempCopyPath, overwrite: true);

                }
                catch (Exception ex)
                {
                    return new
                    {
                        success = false,
                        error = $"Не удалось создать копию файла журнала. Возможно, файл слишком большой или нет прав доступа. {ex.Message}"
                    };
                }

                // 7. Читаем содержимое из копии с правильной кодировкой
                string journalContent;
                try
                {
                    // Журналы Revit в русской Windows обычно в кодировке Windows-1251
                    journalContent = File.ReadAllText(tempCopyPath, Encoding.GetEncoding("windows-1251"));
                }
                catch
                {
                    // Fallback: системная ANSI кодировка
                    journalContent = File.ReadAllText(tempCopyPath, Encoding.Default);
                }

                // 8. Удаляем временную копию
                try
                {
                    File.Delete(tempCopyPath);
                }
                catch { /* Игнорируем ошибки при удалении */ }

                // 9. Извлекаем записи начиная с указанной даты
                var extractedEntries = ExtractEntriesFromDate(journalContent, targetDateTime);

                return new
                {
                    success = true,
                    journal_file = latestJournal.FullName,
                    journal_date = latestJournal.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    target_date = targetDateTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    entries = extractedEntries.entries,
                    entry_count = extractedEntries.count,
                    total_size_bytes = extractedEntries.totalSizeBytes,
                    debug_info = extractedEntries.debugInfo,
                    message = extractedEntries.count == 0
                        ? $"Записи после {targetDateTime:yyyy-MM-dd HH:mm:ss} не найдены"
                        : null
                };
            }
            catch (Exception ex)
            {
                return new
                {
                    success = false,
                    error = $"Ошибка при чтении журнала: {ex.Message}"
                };
            }
        }

        /// <summary>
        /// Парсит пользовательский ввод даты/времени
        /// </summary>
        private static DateTime ParseUserDateTime(string userInput)
        {
            if (string.IsNullOrWhiteSpace(userInput))
                return DateTime.MinValue;

            userInput = userInput.Trim();
            int currentYear = DateTime.Now.Year;

            // Попытка распознать "26.03.2026" или "26.03.2026 14:30"
            if (DateTime.TryParse(userInput, out DateTime parsed))
            {
                return parsed;
            }


            // Попытка распознать "26 марта" или "26 марта 2026"
            var monthMatch = Regex.Match(userInput, @"(\d+)\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)(?:\s+(\d+))?");
            if (monthMatch.Success)
            {
                int day = int.Parse(monthMatch.Groups[1].Value);
                string monthName = monthMatch.Groups[2].Value;
                int year = monthMatch.Groups[3].Success ? int.Parse(monthMatch.Groups[3].Value) : currentYear;
                int month = GetMonthNumberFromRussian(monthName);

                if (month > 0 && day >= 1 && day <= 31)
                {
                    return new DateTime(year, month, day, 0, 0, 0);
                }
            }


            // Попытка распознать "26.03" (без года)
            var shortDateMatch = Regex.Match(userInput, @"(\d+)\.(\d+)");
            if (shortDateMatch.Success)
            {
                int day = int.Parse(shortDateMatch.Groups[1].Value);
                int month = int.Parse(shortDateMatch.Groups[2].Value);
                if (month >= 1 && month <= 12 && day >= 1 && day <= 31)
                {
                    return new DateTime(currentYear, month, day, 0, 0, 0);
                }
            }

            return DateTime.MinValue;
        }

        /// <summary>
        /// Получает номер месяца из русского названия
        /// </summary>
        private static int GetMonthNumberFromRussian(string monthName)
        {
            var months = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "января", 1 }, { "февраля", 2 }, { "марта", 3 }, { "апреля", 4 },
                { "мая", 5 }, { "июня", 6 }, { "июля", 7 }, { "августа", 8 },
                { "сентября", 9 }, { "октября", 10 }, { "ноября", 11 }, { "декабря", 12 }
            };
            return months.ContainsKey(monthName) ? months[monthName] : 0;
        }

        /// <summary>
        /// Формирует путь к папке журналов Revit
        /// </summary>
        private static string GetJournalFolderPath(string revitVersion)
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(localAppData, "Autodesk", "Revit", revitVersion, "Journals");
        }

        /// <summary>
        /// Извлекает записи из журнала начиная с указанной даты
        /// </summary>
        private static (string entries, int count, long totalSizeBytes, string debugInfo) ExtractEntriesFromDate(string journalContent, DateTime targetDate)
        {
            var lines = journalContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var resultLines = new List<string>();
            bool startCollecting = false;
            int entryCount = 0;
            long totalSize = 0;

            // Переменные для отладки
            int totalLinesWithDate = 0;
            string firstDateLine = null;
            string lastDateLine = null;
            DateTime? lastFoundDate = null;
            DateTime? firstFoundDate = null;

            foreach (string line in lines)
            {
                // Пропускаем пустые строки
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Ищем метку времени в формате: "'C 08-Jun-2026 19:22:47.285;"
                // Вариант 1: с буквой после апострофа ('C)
                var timeMatch = Regex.Match(line, @"^'\w\s+(\d{2})-(\w{3})-(\d{4})\s+(\d{2}):(\d{2}):(\d{2})\.\d+;");

                // Вариант 2: без буквы после апострофа (' 08-Jun-2026...)
                if (!timeMatch.Success)
                {
                    timeMatch = Regex.Match(line, @"^'\s+(\d{2})-(\w{3})-(\d{4})\s+(\d{2}):(\d{2}):(\d{2})\.\d+;");
                }

                if (timeMatch.Success)
                {
                    totalLinesWithDate++;

                    int day = int.Parse(timeMatch.Groups[1].Value);
                    string monthAbbr = timeMatch.Groups[2].Value;
                    int year = int.Parse(timeMatch.Groups[3].Value);
                    int hour = int.Parse(timeMatch.Groups[4].Value);
                    int minute = int.Parse(timeMatch.Groups[5].Value);
                    int second = int.Parse(timeMatch.Groups[6].Value);

                    int month = GetMonthNumberFromAbbreviation(monthAbbr);
                    if (month > 0)
                    {
                        DateTime lineDate = new DateTime(year, month, day, hour, minute, second);

                        if (firstFoundDate == null) firstFoundDate = lineDate;
                        lastFoundDate = lineDate;

                        if (!startCollecting && lineDate >= targetDate)
                        {
                            startCollecting = true;
                        }

                        if (startCollecting)
                        {
                            entryCount++;
                            if (firstDateLine == null) firstDateLine = line;
                            lastDateLine = line;
                        }
                    }
                }

                if (startCollecting)
                {
                    resultLines.Add(line);
                    totalSize += line.Length + 2;
                }
            }

            string entries = string.Join("\r\n", resultLines);

            // Формируем отладочную информацию
            string debugInfo = $"Всего строк с датами: {totalLinesWithDate}; " +
                               $"Первая дата в журнале: {firstFoundDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "не найдена"}; " +
                               $"Последняя дата в журнале: {lastFoundDate?.ToString("yyyy-MM-dd HH:mm:ss") ?? "не найдена"}; " +
                               $"Целевая дата: {targetDate:yyyy-MM-dd HH:mm:ss}; " +
                               $"Сбор начат: {startCollecting}; " +
                               $"Найдено записей: {entryCount}";

            // Если размер слишком большой (>200KB), обрезаем с предупреждением
            if (totalSize > 200 * 1024)
            {
                int originalLength = entries.Length;
                entries = entries.Substring(0, 200 * 1024) +
                    $"\r\n\r\n... [ОБРЕЗАНО] Журнал слишком большой. Показаны первые 200 КБ из {totalSize / 1024} КБ";
            }

            return (entries, entryCount, totalSize, debugInfo);
        }


        /// <summary>
        /// Получает номер месяца из английской аббревиатуры (Jun -> 6)
        /// </summary>
        private static int GetMonthNumberFromAbbreviation(string monthAbbr)
        {
            var months = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "Jan", 1 }, { "Feb", 2 }, { "Mar", 3 }, { "Apr", 4 },
                { "May", 5 }, { "Jun", 6 }, { "Jul", 7 }, { "Aug", 8 },
                { "Sep", 9 }, { "Oct", 10 }, { "Nov", 11 }, { "Dec", 12 }
            };
            return months.ContainsKey(monthAbbr) ? months[monthAbbr] : 0;
        }

        /// <summary>
        /// Получает номер месяца из короткого названия ('май' -> 5)
        /// </summary>
        private static int GetMonthNumberFromShort(string monthShort)
        {
            var months = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                { "янв", 1 }, { "фев", 2 }, { "мар", 3 }, { "апр", 4 },
                { "май", 5 }, { "июн", 6 }, { "июл", 7 }, { "авг", 8 },
                { "сен", 9 }, { "окт", 10 }, { "ноя", 11 }, { "дек", 12 }
            };
            return months.ContainsKey(monthShort) ? months[monthShort] : 0;
        }

        #endregion

        #region 99_get_document_switched

        //public static object GetDocumentSwitched(Document mainDoc, UIDocument uiDoc, int elementId = -1, bool switchMainDoc = false)
        //{
        //    try
        //    {
        //        // Случай 1: Возврат к основному документу
        //        if (switchMainDoc)
        //        {
        //            if (_mainDoc == null)
        //            {
        //                // Если сохранённого основного документа нет, используем переданный
        //                _mainDoc = mainDoc;
        //            }

        //            _currentLinkedDoc = null;

        //            return new
        //            {
        //                success = true,
        //                current_document = _mainDoc.Title,
        //                language_of_model = GetDocumentLanguage(_mainDoc),
        //                is_linked = false,
        //                message = $"Переключён на основной документ: {_mainDoc.Title}"
        //            };
        //        }

        //        // Случай 2: Переключение на связанный документ
        //        if (elementId == -1 || elementId == 0)
        //        {
        //            return new
        //            {
        //                success = false,
        //                error = "Не указан elementId RevitLinkInstance для переключения на связанный документ",
        //                current_document = mainDoc.Title,
        //                language_of_model = GetDocumentLanguage(mainDoc)
        //            };
        //        }

        //        ElementId linkElemId = IDHelper.ToElementId(elementId);
        //        Element element = mainDoc.GetElement(linkElemId);

        //        if (element == null)
        //        {
        //            return new
        //            {
        //                success = false,
        //                error = $"Элемент с ID {elementId} не найден в документе",
        //                current_document = mainDoc.Title,
        //                language_of_model = GetDocumentLanguage(mainDoc)
        //            };
        //        }

        //        // Проверяем, является ли элемент RevitLinkInstance
        //        RevitLinkInstance linkInstance = element as RevitLinkInstance;
        //        if (linkInstance == null)
        //        {
        //            return new
        //            {
        //                success = false,
        //                error = $"Элемент с ID {elementId} не является RevitLinkInstance. Укажите ID связанного файла.",
        //                current_document = mainDoc.Title,
        //                language_of_model = GetDocumentLanguage(mainDoc)
        //            };
        //        }




        //        // Получаем связанный документ
        //        Document linkedDoc = linkInstance.GetLinkDocument();

        //        if (linkedDoc == null)
        //        {
        //            return new
        //            {
        //                success = false,
        //                error = "Не удалось получить связанный документ. Возможно, ссылка не загружена или повреждена.",
        //                current_document = mainDoc.Title,
        //                language_of_model = GetDocumentLanguage(mainDoc)
        //            };
        //        }

        //        // Сохраняем основной документ, если ещё не сохранён
        //        if (_mainDoc == null)
        //        {
        //            _mainDoc = mainDoc;
        //        }

        //        // Переключаем контекст
        //        _currentLinkedDoc = linkedDoc;

        //        // Получаем информацию о связанном документе
        //        string linkedDocTitle = linkedDoc.Title ?? "Unnamed";
        //        string language = GetDocumentLanguage(linkedDoc);

        //        // Дополнительная информация о связи
        //        string linkPath = linkInstance.GetLinkDocument().PathName ?? "Unknown";
        //        Transform transform = linkInstance.GetTotalTransform();

        //        return new
        //        {
        //            success = true,
        //            current_document = linkedDocTitle,
        //            language_of_model = language,
        //            is_linked = true,
        //            link_info = new
        //            {
        //                link_instance_id = elementId,
        //                link_file_name = linkPath,                       
        //                transform_translation = new
        //                {
        //                    x = transform.Origin.X,
        //                    y = transform.Origin.Y,
        //                    z = transform.Origin.Z
        //                }
        //            },
        //            message = $"Переключён на связанный документ: {linkedDocTitle}"
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new
        //        {
        //            success = false,
        //            error = $"Ошибка при переключении документа: {ex.Message}",
        //            current_document = mainDoc?.Title ?? "Unknown",
        //            language_of_model = GetDocumentLanguage(mainDoc)
        //        };
        //    }
        //}

        /// <summary>
        /// Получает текущий активный документ (основной или связанный)
        /// </summary>
        //public static Document GetCurrentDocument()
        //{
        //    return _currentLinkedDoc ?? _mainDoc;
        //}

        /// <summary>
        /// Сбрасывает контекст к основному документу
        /// </summary>
        //public static void ResetDocumentContext()
        //{
        //    _currentLinkedDoc = null;
        //    // _mainDoc не сбрасываем, чтобы сохранить ссылку
        //}

        /// <summary>
        /// Получает язык интерфейса ревит
        /// </summary>
        //private static string GetDocumentLanguage(Document doc)
        //{
        //    try
        //    {
        //        // Получаем язык через Application
        //        Application app = doc.Application;
        //        LanguageType language = app.Language;

        //        switch (language)
        //        {
        //            case LanguageType.English_USA:
        //                return "English (US)";
        //            case LanguageType.Russian:
        //                return "Russian (Русский)";
        //            case LanguageType.German:
        //                return "German (Deutsch)";
        //            case LanguageType.French:
        //                return "French (Français)";
        //            case LanguageType.Spanish:
        //                return "Spanish (Español)";
        //            case LanguageType.Italian:
        //                return "Italian (Italiano)";
        //            case LanguageType.Dutch:
        //                return "Dutch (Nederlands)";
        //            case LanguageType.Chinese_Simplified:
        //                return "Chinese Simplified (简体中文)";
        //            case LanguageType.Chinese_Traditional:
        //                return "Chinese Traditional (繁體中文)";
        //            case LanguageType.Japanese:
        //                return "Japanese (日本語)";
        //            case LanguageType.Korean:
        //                return "Korean (한국어)";
        //            default:
        //                return language.ToString();
        //        }
        //    }
        //    catch
        //    {
        //        return "Unknown";
        //    }
        //}


        #endregion


    }
}