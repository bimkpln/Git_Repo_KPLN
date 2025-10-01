using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib.Common;
using System.Collections.Generic;
using System.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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

        /// <summary>
        /// Конструктор для генерации замечания о отсутсвии элементов в файле
        /// </summary>
        public CheckerEntity(string header, string description, string info)
        {
            Status = ErrorStatus.Error;
            
            Header = header;
            Description = description;
            Info = info;
        }

        /// <summary>
        /// Основной конструктор
        /// </summary>
        public CheckerEntity(
            object elemData, 
            string header, 
            string description, 
            string info, 
            bool canZoomed = false)
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
        public bool CanZoomed { get; private set; }

        /// <summary>
        /// Есть возможность подтверждать ошибку?
        /// </summary>
        public bool CanApproved { get; private set; }

        /// <summary>
        /// Комментарий, указанный при подтверждении
        /// </summary>
        public string ApproveComment { get; private set; }

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
            private set => _zoomBBox = value;
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
            private set => _zoomCentroid = value;
        }

        /// <summary>
        /// Взвести метку, что замечание может быть подтверждено юзером, а также установить данные из ExtStorage (статус и комментарий)
        /// </summary>
        /// <param name="esEntity">ExtensibleStorageEntity проверки</param>
        /// <param name="statusIfNotApprove">Статус, который нужно присвоить замечанию, если оно не помечено юзером</param>
        /// <returns></returns>
        public CheckerEntity Set_CanApprovedAndESData(ExtensibleStorageEntity esEntity, ErrorStatus statusIfNotApprove = ErrorStatus.Error)
        {
            CanApproved = true;

            if (Element != null)
            {
                if (esEntity.ESBuilderUserText.IsDataExists_Text(Element))
                {
                    Status = ErrorStatus.Approve;
                    ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(Element).Description;
                }
                else
                    Status = statusIfNotApprove;
            }
            else if (ElementCollection != null && ElementCollection.Any())
            {
                if (ElementCollection.All(e => esEntity.ESBuilderUserText.IsDataExists_Text(e))
                    && ElementCollection.All(e =>
                        esEntity.ESBuilderUserText.GetResMessage_Element(e).Description
                            .Equals(esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description)))
                {
                    Status = ErrorStatus.Approve;
                    ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description;
                }
                else
                    Status = statusIfNotApprove;
            }

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
        /// Вручную установить ZoomBBox и ZoomCentroid.
        /// </summary>
        /// <param name="box">BoundingBoxXYZ для установки</param>
        /// <returns></returns>
        public CheckerEntity Set_ZoomData(BoundingBoxXYZ box)
        {
            if (box != null)
            {
                ZoomBBox = box;
                ZoomCentroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
            }

            return this;
        }
    }
}
