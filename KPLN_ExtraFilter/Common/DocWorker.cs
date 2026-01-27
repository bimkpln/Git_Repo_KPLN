using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ExtraFilter.Common
{
    /// <summary>
    /// Класс для сравнения категорий
    /// </summary>
    internal class CategoryComparer : IEqualityComparer<Category>
    {
        public bool Equals(Category x, Category y) => x.Id == x.Id;

        public int GetHashCode(Category obj) => obj.Id.GetHashCode();
    }

    /// <summary>
    /// Класс для сравнения парамтеров
    /// </summary>
    internal class ParameterComparer : IEqualityComparer<Parameter>
    {
        public bool Equals(Parameter x, Parameter y) => x.Definition.Name == y.Definition.Name;

        public int GetHashCode(Parameter obj) => obj.Definition.Name.GetHashCode();
    }

    internal sealed class DocWorker
    {
        /// <summary>
        /// Получить ВСЕ категории из указанных элементов
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="elemsToFind">Коллекция элементов для анализа</param>
        /// <returns></returns>
        internal static IEnumerable<Category> GetAllCatsFromElems(IEnumerable<Element> elemsToFind)
        {
            if (elemsToFind.All(el => el == null))
                return null;

            HashSet<Category> docCats = new HashSet<Category>(new CategoryComparer());
            foreach (Element currentElement in elemsToFind)
            {
                docCats.Add(currentElement.Category);
            }

            return docCats;
        }

        /// <summary>
        /// Получить ВООБЩЕ ВСЕ парамеры из указанных элементов
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="elemsToFind">Коллекция элементов для анализа</param>
        /// <returns></returns>
        internal static IEnumerable<Parameter> GetAllParamsFromElems(Document doc, IEnumerable<Element> elemsToFind, bool exceptReadOnly)
        {
            if (elemsToFind.All(el => el == null))
                return null;

            IEnumerable<Element> oneTypeIDElColl = FilterElemCollByTypeId(elemsToFind);


            // Основной блок
            HashSet<Parameter> commonInstParameters = new HashSet<Parameter>(new ParameterComparer());
            HashSet<Parameter> commonTypeParameters = new HashSet<Parameter>(new ParameterComparer());
            Element firstElement = oneTypeIDElColl.FirstOrDefault();
            AddParam(firstElement, commonInstParameters, exceptReadOnly);

            // Не у всех элементов есть возможность выбрать тип. Например - помещения
            if (doc.GetElement(firstElement.GetTypeId()) is Element typeElem)
                AddParam(typeElem, commonTypeParameters, exceptReadOnly);

            int countElemsToFind = oneTypeIDElColl.Count();
            foreach (Element currentElement in oneTypeIDElColl)
            {
                // Игнорирую уже добавленный эл-т
                if (countElemsToFind > 1
                    && firstElement.Id == currentElement.Id)
                    continue;

                HashSet<Parameter> currentInstParameters = new HashSet<Parameter>(new ParameterComparer());
                AddParam(currentElement, currentInstParameters, exceptReadOnly);
                commonInstParameters.UnionWith(currentInstParameters);

                // Не у всех элементов есть возможность выбрать тип. Например - помещения
                if (doc.GetElement(firstElement.GetTypeId()) is Element currentTypeElem)
                {
                    HashSet<Parameter> currentTypeParameters = new HashSet<Parameter>(new ParameterComparer());
                    AddParam(currentTypeElem, currentTypeParameters, exceptReadOnly);
                    commonTypeParameters.UnionWith(currentTypeParameters);
                }
            }

            if (commonInstParameters.Count == 0 && commonTypeParameters.Count == 0)
                return null;

            return new HashSet<Parameter>(commonInstParameters.Union(commonTypeParameters),
                new ParameterComparer());
        }


        /// <summary>
        /// Получить ОБЩИЕ парамеры из указанных элементов
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="elemsToFind">Коллекция элементов для анализа</param>
        /// <returns></returns>
        internal static IEnumerable<Parameter> GetUnionParamsFromElems(Document doc, Element[] elemsToFind, bool exceptReadOnly)
        {
            if (elemsToFind.All(el => el == null))
                return null;

            IEnumerable<Element> oneTypeIDElColl = FilterElemCollByTypeId(elemsToFind);


            // Основной блок
            HashSet<Parameter> commonInstParameters = new HashSet<Parameter>(new ParameterComparer());
            HashSet<Parameter> commonTypeParameters = new HashSet<Parameter>(new ParameterComparer());
            Element firstElement = oneTypeIDElColl.FirstOrDefault();
            AddParam(firstElement, commonInstParameters, exceptReadOnly);

            // Не у всех элементов есть возможность выбрать тип. Например - помещения
            if (doc.GetElement(firstElement.GetTypeId()) is Element typeElem)
                AddParam(typeElem, commonTypeParameters, exceptReadOnly);

            int countElemsToFind = oneTypeIDElColl.Count();
            foreach (Element currentElement in oneTypeIDElColl)
            {
                // Игнорирую уже добавленный эл-т
                if (countElemsToFind > 1
                    && firstElement.Id == currentElement.Id)
                    continue;

                HashSet<Parameter> currentInstParameters = new HashSet<Parameter>(new ParameterComparer());
                AddParam(currentElement, currentInstParameters, exceptReadOnly);
                commonInstParameters.IntersectWith(currentInstParameters);

                // Не у всех элементов есть возможность выбрать тип. Например - помещения
                if (doc.GetElement(firstElement.GetTypeId()) is Element currentTypeElem)
                {
                    HashSet<Parameter> currentTypeParameters = new HashSet<Parameter>(new ParameterComparer());
                    AddParam(currentTypeElem, currentTypeParameters, exceptReadOnly);
                    commonTypeParameters.IntersectWith(currentTypeParameters);
                }
            }

            if (commonInstParameters.Count == 0 && commonTypeParameters.Count == 0)
                return null;

            return new HashSet<Parameter>(commonInstParameters.Union(commonTypeParameters),
                new ParameterComparer());
        }

        /// <summary>
        /// Перевод единиц проекта с систему СИ
        /// </summary>
        /// <param name="userSelParam"></param>
        /// <param name="doc"></param>
        /// <returns></returns>
        public static string GetParamValueInSI(Document doc, Parameter userSelParam)
        {
            string paramData = null;

            switch (userSelParam.StorageType)
            {
                case StorageType.ElementId:
                    paramData = userSelParam.AsValueString();
                    break;

                case StorageType.String:
                    paramData = userSelParam.AsString();
                    break;

                case StorageType.Integer:
                    paramData = userSelParam.AsInteger().ToString();
                    break;

                case StorageType.Double:
                    double value = userSelParam.AsDouble();

#if Debug2020 || Revit2020
                    DisplayUnitType displayUnit = userSelParam.DisplayUnitType;
                    double siValue = UnitUtils.ConvertFromInternalUnits(value, displayUnit);
                    paramData = $"{siValue:0.###}";

#else
                    // Атрымліваем адзінку вымярэння параметра
                    ForgeTypeId unitTypeId = userSelParam.GetUnitTypeId();

                    // Канвертацыя з футаў у СІ
                    double siValue = UnitUtils.ConvertFromInternalUnits(value, unitTypeId);

                    // Афармленне з 3 знакамі пасля коскі
                    paramData = $"{siValue:0.###}";
#endif
                    break;
            }

            return paramData;
        }

        /// <summary>
        /// Добавить пар-р в коллекцию с пред. подготовкой
        /// </summary>
        /// <param name="elem">Елемент для аналища</param>
        /// <param name="setToAdd">Коллекция для добавления</param>
        private static void AddParam(Element elem, HashSet<Parameter> setToAdd, bool exceptReadOnly)
        {
            if (elem == null)
                return;

            foreach (Parameter param in elem.Parameters)
            {
                if (exceptReadOnly && param.IsReadOnly)
                    continue;
                
                if (param.Definition == null)
                    continue;

                StorageType paramST = param.StorageType;
                string paramNameLC = param.Definition.Name.ToLower();

                // Отбрасываю лишние пара-ры
                if (paramST == StorageType.None
                    || paramNameLC.Contains("ifc")
                    || paramNameLC.Contains("url"))
                    continue;

                setToAdd.Add(param);
            }
        }

        /// <summary>
        /// Отфильтровать коллекцию, чтобы взять по ОДНОМУ экземпляру КАЖДОГО типа
        /// </summary>
        /// <param name="elemsToFind"></param>
        /// <returns></returns>
        private static IEnumerable<Element> FilterElemCollByTypeId(IEnumerable<Element> elemsToFind) =>
            elemsToFind
                .GroupBy(el => el.GetTypeId())
                .Select(g => g.First());
    }
}
