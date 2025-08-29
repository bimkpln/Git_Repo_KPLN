using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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

        public CheckerEntity(object elemData, string header, string description, string info, bool isZoomElement)
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
                ElementCollection = elements;
                ElementIdCollection = elements.Select(e => e.Id);

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
            IsZoomElement = isZoomElement;
        }

        public CheckerEntity(object elemData, string header, string description, string info, bool isZoomElement, ErrorStatus status) : this(elemData, header, description, info, isZoomElement)
        {
            Status = status;
        }

        /// <summary>
        /// Revit-элемент
        /// </summary>
        public Element Element { get; }

        /// <summary>
        /// Коллекция Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public IEnumerable<Element> ElementCollection { get; } = Enumerable.Empty<Element>();

        /// <summary>
        /// Коллекция Id Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public IEnumerable<ElementId> ElementIdCollection { get; } = Enumerable.Empty<ElementId>();

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
        public ErrorStatus Status { get; } = ErrorStatus.Error;

        /// <summary>
        /// Можно использовать кастомный зум?
        /// </summary>
        public bool IsZoomElement { get; }

        /// <summary>
        /// BoundingBoxXYZ для зума
        /// </summary>
        public BoundingBoxXYZ ZoomBBox
        {
            get
            {
                if (IsZoomElement && _zoomBBox == null)
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
                if (IsZoomElement && _zoomCentroid == null)
                {
                    if (ZoomBBox != null)
                        _zoomCentroid = new XYZ(
                            (ZoomBBox.Min.X + ZoomBBox.Max.X) / 2,
                            (ZoomBBox.Min.Y + ZoomBBox.Max.Y) / 2,
                            (ZoomBBox.Min.Z + ZoomBBox.Max.Z) / 2);
                    else
                    {
                        LocationPoint locationPoint = Element.Location as LocationPoint;
                        _zoomCentroid = locationPoint.Point;
                    }
                }

                return _zoomCentroid;
            }
        }
    }
}
