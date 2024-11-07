﻿using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ExtraFilter.Common
{
    /// <summary>
    /// Класс для сравнения парамтеров
    /// </summary>
    internal class ParameterComparer : IEqualityComparer<Parameter>
    {
        public bool Equals(Parameter x, Parameter y) => x.Definition.Name == y.Definition.Name;

        public int GetHashCode(Parameter obj) => obj.Definition.Name.GetHashCode();
    }

    internal class ParamWorker
    {
        /// <summary>
        /// Получить парамеры из указанных элементов
        /// </summary>
        /// <param name="doc">Ревит-файл</param>
        /// <param name="elemsToFind">Коллекция элементов для анализа</param>
        /// <returns></returns>
        internal static IEnumerable<Parameter> GetParamsFromElems(Document doc, Element[] elemsToFind)
        {
            // Основной блок
            HashSet<Parameter> commonInstParameters = new HashSet<Parameter>(new ParameterComparer());
            HashSet<Parameter> commonTypeParameters = new HashSet<Parameter>(new ParameterComparer());
            Element firstElement = elemsToFind.FirstOrDefault();

            AddParam(firstElement, commonInstParameters);

            Element typeElem = doc.GetElement(firstElement.GetTypeId());
            AddParam(typeElem, commonTypeParameters);

            foreach (Element currentElement in elemsToFind)
            {
                // Игнорирую уже добавленный эл-т
                if (firstElement.Id == currentElement.Id)
                    continue;

                HashSet<Parameter> currentInstParameters = new HashSet<Parameter>(new ParameterComparer());
                AddParam(currentElement, currentInstParameters);
                commonInstParameters.IntersectWith(currentInstParameters);

                HashSet<Parameter> currentTypeParameters = new HashSet<Parameter>(new ParameterComparer());
                Element currentTypeElem = doc.GetElement(currentElement.GetTypeId());
                AddParam(currentTypeElem, currentTypeParameters);
                commonTypeParameters.IntersectWith(currentTypeParameters);
            }

            if (commonInstParameters.Count == 0 && commonTypeParameters.Count == 0)
                throw new Exception("Ошибка в поиске парамеров для элементов Revit");

            return new HashSet<Parameter>(commonInstParameters.Union(commonTypeParameters),
                new ParameterComparer());
        }

        /// <summary>
        /// Добавить пар-р в коллекцию с пред. подготовкой
        /// </summary>
        /// <param name="elem">Елемент для аналища</param>
        /// <param name="setToAdd">Коллекция для добавления</param>
        private static void AddParam(Element elem, HashSet<Parameter> setToAdd)
        {
            foreach (Parameter param in elem.Parameters)
            {
                // Отбрасываю системные пар-ры, которые нельзя редачить (Категория, Имя типа и т.п.) 
                if (param.Id.IntegerValue < 0 && param.IsReadOnly)
                    continue;

                setToAdd.Add(param);
            }
        }
    }
}
