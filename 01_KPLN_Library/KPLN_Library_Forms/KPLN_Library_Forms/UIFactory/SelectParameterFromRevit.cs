using Autodesk.Revit.DB;
using KPLN_Library_Forms.Common;
using KPLN_Library_Forms.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;

namespace KPLN_Library_Forms.UIFactory
{
    /// <summary>
    /// Выбор параметра из проекта
    /// </summary>
    public static class SelectParameterFromRevit
    {
        /// <summary>
        /// Запуск окна выбора параметра по коллекции элементов
        /// </summary>
        /// <param name="doc">Проект ревит</param>
        /// <param name="elemsToFindParam">Коллекция для анализа</param>
        /// <param name="storageType">StorageType - тип параметров, которые нужно учесть</param>
        /// <param name="includeIsntParams">Учитывать пар-ры экземпляров</param>
        /// <param name="includeTypeParams">Учитывать пар-ры типов</param>
        /// <returns></returns>
        public static ElementSinglePick CreateForm(
            Window owner,
            Document doc,
            IEnumerable<Element> elemsToFindParam,
            StorageType storageType,
            bool includeIsntParams = true,
            bool includeTypeParams = true)
        {
            // Чистка коллекции от экз. одинаковых семейсвт (многопоточность не справиться из-за ревит, поэтому нужно предв. очистка)
            IEnumerable<Element> clearedElemsToFind = elemsToFindParam
                .GroupBy(x => x.GetTypeId())
                .Select(gr => gr.FirstOrDefault());

            // Основной блок
            HashSet<Parameter> commonInstParameters = new HashSet<Parameter>(new ParameterComparer());
            HashSet<Parameter> commonTypeParameters = new HashSet<Parameter>(new ParameterComparer());
            Element firstElement = clearedElemsToFind.FirstOrDefault();

            if (includeIsntParams)
                AddParam(firstElement, storageType, commonInstParameters);

            if (includeTypeParams)
            {
                Element typeElem = doc.GetElement(firstElement.GetTypeId());
                AddParam(typeElem, storageType, commonTypeParameters);
            }

            foreach (Element currentElement in clearedElemsToFind)
            {
                // Игнорирую уже добавленный эл-т
                if (firstElement.Id == currentElement.Id)
                    continue;

                HashSet<Parameter> currentInstParameters = new HashSet<Parameter>(new ParameterComparer());
                HashSet<Parameter> currentTypeParameters = new HashSet<Parameter>(new ParameterComparer());
                if (includeIsntParams)
                {
                    AddParam(currentElement, storageType, currentInstParameters);
                    commonInstParameters.IntersectWith(currentInstParameters);
                }

                if (includeTypeParams)
                {
                    Element typeElem = doc.GetElement(currentElement.GetTypeId());
                    AddParam(typeElem, storageType, currentTypeParameters);
                    commonTypeParameters.IntersectWith(currentTypeParameters);
                }
            }

            if (commonInstParameters.Count == 0 && commonTypeParameters.Count == 0)
                throw new Exception("Ошибка в поиске парамтеров для элементов Revit");

            ObservableCollection<ElementEntity> paramsToChoose = new ObservableCollection<ElementEntity>();
            HashSet<Parameter> resultParamsColl = new HashSet<Parameter>(commonInstParameters.Union(commonTypeParameters), new ParameterComparer());
            foreach (Parameter param in resultParamsColl)
            {
                string toolTip = string.Empty;
                if (param.IsShared)
                    toolTip = $"Id: {param.Id}, GUID: {param.GUID}";
#if !Debug2020 && !Debug2023 && !Revit2020 && !Revit2023
                else if (param.Id.Value < 0)
                    toolTip = $"Id: {param.Id}, это СИСТЕМНЫЙ параметр проекта";
#else
                else if (param.Id.IntegerValue < 0)
                    toolTip = $"Id: {param.Id}, это СИСТЕМНЫЙ параметр проекта";
#endif
                else
                    toolTip = $"Id: {param.Id}, это ПОЛЬЗОВАТЕЛЬСКИЙ параметр проекта";

                paramsToChoose.Add(new ElementEntity(param, toolTip));
            }

            ElementSinglePick pickForm = new ElementSinglePick(owner, paramsToChoose.OrderBy(p => p.Name), "Выбери пар-р, общий для необходимых категорий проекта");

            return pickForm;
        }

        private static void AddParam(Element elem, StorageType storageType, HashSet<Parameter> setToAdd)
        {
            foreach (Parameter param in elem.Parameters)
            {
                if (param.StorageType == storageType)
                    setToAdd.Add(param);
            }
        }
    }

    /// <summary>
    /// Класс для сравнения парамтеров
    /// </summary>
    public class ParameterComparer : IEqualityComparer<Parameter>
    {
        public bool Equals(Parameter x, Parameter y) => x.Definition.Name == y.Definition.Name;

        public int GetHashCode(Parameter obj) => obj.Definition.Name.GetHashCode();
    }
}
