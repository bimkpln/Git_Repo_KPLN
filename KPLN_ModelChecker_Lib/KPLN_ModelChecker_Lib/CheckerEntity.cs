using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib
{
    /// <summary>
    /// Основная коллекция статусов ошибок
    /// </summary>
    public enum ErrorStatus
    {
        Error,
        Warning,
        LittleWarning,
        AllmostOk,
        Approve,
    }

    /// <summary>
    /// Сущность для аккумулирования данных по ошибкам
    /// </summary>
    public sealed class CheckerEntity
    {
        private BoundingBoxXYZ _zoomBBox;
        private XYZ _zoomCentroid;

        public CheckerEntity(object elemData, string header, string description, string info, bool canZoomed)
        {
            if (elemData is Element element)
            {
                Element = element;
                ElementIdCollection = new ElementId[] { Element.Id };

                if (Element is Room room)
                    ElementName = room.Name;
                else
                    ElementName = !(element is FamilyInstance familyInstance) ? element.Name : $"{familyInstance.Symbol.FamilyName}: {element.Name}";
                if (Element is Family family)
                    CategoryName = family.FamilyCategory.Name;
                else if (Element is ElementType elType)
                    CategoryName = elType.FamilyName;
                else
                    CategoryName = element.Category.Name;
            }
            else if (elemData is IEnumerable<Element> elements)
            {
                ElementCollection = elements.ToArray();
                ElementIdCollection = elements
                    .Where(e => e.IsValidObject)
                    .Select(e => e.Id)
                    .ToArray();

                if (ElementCollection.Count() > 1)
                {
                    ElementName = "<Набор элементов>";
                    HashSet<string> uniqueElemCatNames = new HashSet<string>(ElementCollection.Select(e => e.Category.Name));
                    if (uniqueElemCatNames.Count() > 1) CategoryName = "<Набор категорий>";
                    else CategoryName = uniqueElemCatNames.FirstOrDefault();
                }
                else if (ElementCollection.Count() == 1)
                {
                    Element currentElem = ElementCollection.FirstOrDefault();
                    ElementName = !(currentElem is FamilyInstance familyInstance) ? currentElem.Name : $"{familyInstance.Symbol.FamilyName}: {currentElem.Name}";
                    CategoryName = currentElem.Category.Name;
                }
            }


            Header = header;
            Description = description;
            Info = info;
            CanZoomed = canZoomed;
        }

        /// <summary>
        /// Взвести метку, что замечание может быть подтверждено юзером
        /// </summary>
        /// <returns></returns>
        public CheckerEntity Set_CanApproved()
        {
            CanApproved = true;
            
            return this;
        }

        /// <summary>
        /// Вручную указать статус замечания
        /// </summary>
        /// <param name="status">Статус, который нужно присвоить замечанию</param>
        /// <returns></returns>
        public CheckerEntity Set_Status(ErrorStatus status)
        {
            Status = status;

            return this;
        }


        /// <summary>
        /// Revit-элемент
        /// </summary>
        public Element Element { get; }

        /// <summary>
        /// Коллекция Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public Element[] ElementCollection { get; }

        /// <summary>
        /// Коллекция Id Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public ElementId[] ElementIdCollection { get; }

        /// <summary>
        /// Заголовок ошибки
        /// </summary>
        public string Header { get; set; }

        /// <summary>
        /// Имя элемента
        /// </summary>
        public string ElementName { get; set; }

        /// <summary>
        /// Имя категории
        /// </summary>
        public string CategoryName { get; }

        /// <summary>
        /// Описание ошибки
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Дополнительная инструкция
        /// </summary>
        public string Info { get; }

        /// <summary>
        /// Статус ошибки
        /// </summary>
        public ErrorStatus Status { get; private set; } = ErrorStatus.Error;

        /// <summary>
        /// Можно использовать кастомный зум?
        /// </summary>
        public bool CanZoomed { get; private set; } = false;

        /// <summary>
        /// Есть возможность подтверждать ошибку?
        /// </summary>
        public bool CanApproved { get; private set; } = false;

        /// <summary>
        /// BoundingBoxXYZ для зума
        /// </summary>
        public BoundingBoxXYZ ZoomBBox
        {
            get
            {
                if (CanZoomed && _zoomBBox == null)
                    _zoomBBox = Element.get_BoundingBox(null);

                return _zoomBBox;
            }
        }

        /// <summary>
        /// Центроид для зума
        /// </summary>
        public XYZ ZoomCentroid
        {
            get
            {
                if (CanZoomed && _zoomCentroid == null)
                {
                    if (ZoomBBox != null)
                        _zoomCentroid = new XYZ(
                            (ZoomBBox.Min.X + ZoomBBox.Max.X) / 2,
                            (ZoomBBox.Min.Y + ZoomBBox.Max.Y) / 2,
                            (ZoomBBox.Min.Z + ZoomBBox.Max.Z) / 2);
                    else
                    {
                        if (Element.Location is LocationPoint locationPoint)
                            _zoomCentroid = locationPoint.Point;
                    }
                }

                return _zoomCentroid;
            }
        }
    }
}
