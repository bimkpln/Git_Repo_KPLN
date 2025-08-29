using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using KPLN_ModelChecker_Lib;
using KPLN_ModelChecker_User.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Xml.Linq;

namespace KPLN_ModelChecker_User.WPFItems
{
    /// <summary>
    /// Спец. класс-обертка, для передачи в WPFReport для генерации окна-отчёта
    /// </summary>
    public sealed class WPFEntity : INotifyPropertyChanged
    {
        /// <summary>
        /// Фон элемента
        /// </summary>
        private SolidColorBrush _background;
        /// <summary>
        /// Комментарий, указанный при подтверждении
        /// </summary>
        private string _approveComment;
        private string _errorDescription;
        /// <summary>
        /// Заголовок элемента
        /// </summary>
        private string _header;
        private ErrorStatus _currentStatus = ErrorStatus.Error;

        public WPFEntity(CheckerEntity checkEntity, ExtensibleStorageEntity esEntity)
        {
            Element = checkEntity.Element;
            ElementCollection = checkEntity.ElementCollection;

            if (ElementCollection.Any())
            {
                ElementIdCollection = ElementCollection.Select(e => e.Id);

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

                if (ElementCollection.All(e => esEntity.ESBuilderUserText.IsDataExists_Text(e))
                    && ElementCollection.All(e =>
                        esEntity
                        .ESBuilderUserText
                        .GetResMessage_Element(e)
                        .Description
                        .Equals(esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description)))
                    {
                        CurrentStatus = ErrorStatus.Approve;
                        ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description;
                    }
                else
                    CurrentStatus = ErrorStatus.Error;
            }
            else
            {
                ElementIdCollection = new ElementId[] { Element.Id };

                if (Element is Room room)
                    ElementName = room.Name;
                else
                    ElementName = !(Element is FamilyInstance familyInstance) ? Element.Name : $"{familyInstance.Symbol.FamilyName}: {Element.Name}";
                if (Element is Family family)
                    CategoryName = family.FamilyCategory.Name;
                else if (Element is ElementType elType)
                    CategoryName = elType.FamilyName;
                else
                    CategoryName = Element.Category.Name;

                if (esEntity.ESBuilderUserText.IsDataExists_Text(Element))
                {
                    CurrentStatus = ErrorStatus.Approve;
                    ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(Element).Description;
                }
                else
                    CurrentStatus = ErrorStatus.Error;
            }

            Header = checkEntity.Header;
            Description = checkEntity.Description;
            Info = checkEntity.Info;
            CurrentStatus = checkEntity.Status;
            IsZoomElement = checkEntity.IsZoomElement;
            Box = checkEntity.ZoomBBox;
            Centroid = checkEntity.ZoomCentroid;
        }

        [Obsolete]
        public WPFEntity(ExtensibleStorageEntity esEntity, Element element, string header, string description, string info, bool isZoomElement)
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

            Header = header;
            Description = description;
            Info = info;
            IsZoomElement = isZoomElement;

            Box = Element.get_BoundingBox(null);
            // У некоторых эл-в нет BBox.
            if (Box != null)
                Centroid = new XYZ((Box.Min.X + Box.Max.X) / 2, (Box.Min.Y + Box.Max.Y) / 2, (Box.Min.Z + Box.Max.Z) / 2);

            if (esEntity.ESBuilderUserText.IsDataExists_Text(element))
            {
                CurrentStatus = ErrorStatus.Approve;
                ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(element).Description;
            }
            else
                CurrentStatus = ErrorStatus.Error;
        }

        [Obsolete]
        public WPFEntity(ExtensibleStorageEntity esEntity, Element element, string header, string description, string info, bool isZoomElement, ErrorStatus status) : this(esEntity, element, header, description, info, isZoomElement)
        {
            if (CurrentStatus != ErrorStatus.Approve)
                CurrentStatus = status;
        }

        [Obsolete]
        public WPFEntity(ExtensibleStorageEntity esEntity, Element element, string header, string description, string info, bool isZoomElement, ErrorStatus status, bool isApproveElement) : this(esEntity, element, header, description, info, isZoomElement, status)
        {
            IsApproveElement = isApproveElement;
        }

        [Obsolete]
        public WPFEntity(ExtensibleStorageEntity esEntity, Element element, string header, string description, string info, bool isZoomElement, bool isApproveElement) : this(esEntity, element, header, description, info, isZoomElement)
        {
            IsApproveElement = isApproveElement;
        }

        [Obsolete]
        public WPFEntity(ExtensibleStorageEntity esEntity, IEnumerable<Element> elements, string header, string description, string info, bool isZoomElement)
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

            Header = header;
            Description = description;
            Info = info;
            IsZoomElement = isZoomElement;

            if (ElementCollection.All(e => esEntity.ESBuilderUserText.IsDataExists_Text(e))
                && ElementCollection.All(e =>
                    esEntity.ESBuilderUserText.GetResMessage_Element(e).Description
                        .Equals(esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description)))
            {
                CurrentStatus = ErrorStatus.Approve;
                ApproveComment = esEntity.ESBuilderUserText.GetResMessage_Element(ElementCollection.FirstOrDefault()).Description;
            }
            else
                CurrentStatus = ErrorStatus.Error;
        }

        /// <summary>
        /// Revit-элемент
        /// </summary>
        public Element Element { get; }

        /// <summary>
        /// Коллекция Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public IEnumerable<Element> ElementCollection { get; }

        /// <summary>
        /// Коллекция Id Revit-элементов, объединенных одной ошибкой
        /// </summary>
        public IEnumerable<ElementId> ElementIdCollection { get; }

        /// <summary>
        /// Изображение поиска в WPF
        /// </summary>
        public string SearchIcon { get; } = "🔍";

        /// <summary>
        /// Изображение для смены статуса в WPF
        /// </summary>
        public string ApproveIcon { get; } = "✔️";

        /// <summary>
        /// Заголовок в окне
        /// </summary>
        public string FormHeader
        {
            get => $"{ErrorDescription}: {Header}";
        }

        /// <summary>
        /// Описание ошибки
        /// </summary>
        public string ErrorDescription
        {
            get => _errorDescription;
            set
            {
                if (value != _errorDescription)
                {
                    _errorDescription = value;
                    OnPropertyChanged(nameof(ErrorDescription));
                    OnPropertyChanged(nameof(FormHeader));
                }
            }
        }

        /// <summary>
        /// Заголовок элемента
        /// </summary>
        public string Header
        {
            get => _header;
            set
            {
                if (value != _header)
                {
                    _header = value;
                    OnPropertyChanged(nameof(Header));
                }
            }
        }

        public string CategoryName { get; }

        /// <summary>
        /// Спец. имя элемента в отчете
        /// </summary>
        public string ElementName { get; set; }

        /// <summary>
        /// Описание элемента
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Дополнительная инструкция
        /// </summary>
        public string Info { get; }

        /// <summary>
        /// Описание для фильтрации (текстовое значение, по которому группируются элементы)
        /// </summary>
        public string FiltrationDescription { get; set; }

        /// <summary>
        /// Комментарий, указанный при подтверждении
        /// </summary>
        public string ApproveComment
        {
            get => _approveComment;
            set
            {
                if (value != _approveComment)
                {
                    _approveComment = value;
                    OnPropertyChanged(nameof(ApproveComment));
                }
            }
        }

        /// <summary>
        /// Текущий статус ошибки
        /// </summary>
        public ErrorStatus CurrentStatus
        {
            get => _currentStatus;
            private set
            {
                _currentStatus = value;
                OnPropertyChanged(nameof(CurrentStatus));

                UpdateMainFieldByStatus();
            }
        }

        /// <summary>
        /// Использовать кастомный зум?
        /// </summary>
        public bool IsZoomElement { get; }

        /// <summary>
        /// Есть возможность подтверждать ошибку?
        /// </summary>
        public bool IsApproveElement { get; } = false;

        /// <summary>
        /// Фон элемента
        /// </summary>
        public SolidColorBrush Background
        {
            get => _background;
            private set
            {
                _background = value;
                OnPropertyChanged(nameof(Background));
            }
        }

        public BoundingBoxXYZ Box { get; private set; }

        public XYZ Centroid { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Обновление основных визуальных разделителей единицы отчета
        /// </summary>
        /// <param name="status"></param>
        public void UpdateMainFieldByStatus()
        {
            switch (_currentStatus)
            {
                case ErrorStatus.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 240, 90));
                    ErrorDescription = "Обрати внимание";
                    break;
                case ErrorStatus.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    ErrorDescription = "Предупреждение";
                    break;
                case ErrorStatus.Error:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    ErrorDescription = "Ошибка";
                    break;
                case ErrorStatus.Approve:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 125, 105, 240));
                    ErrorDescription = "Допустимое";
                    break;
            }
        }

        /// <summary>
        /// Обновление основных визуальных разделителей единицы отчета
        /// </summary>
        /// <param name="status"></param>
        public void UpdateMainFieldByStatus(ErrorStatus status)
        {
            CurrentStatus = status;
            switch (status)
            {
                case ErrorStatus.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 240, 90));
                    ErrorDescription = "Обрати внимание";
                    break;
                case ErrorStatus.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    ErrorDescription = "Предупреждение";
                    break;
                case ErrorStatus.Error:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    ErrorDescription = "Ошибка";
                    break;
                case ErrorStatus.Approve:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 125, 105, 240));
                    ErrorDescription = "Допустимое";
                    break;
            }
        }

        /// <summary>
        /// Пересоздаёт рамку для зумирования (дефолтная - из конструктора)
        /// </summary>
        public void ResetZoomGeometryExtension(BoundingBoxXYZ box)
        {
            if (box != null)
            {
                Box = box;
                Centroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
            }
        }

        /// <summary>
        /// Осветляет фон элемента
        /// </summary>
        public void BackgroundLightening()
        {
            double lightenPercentage = 0.20;
            System.Windows.Media.Color lightenedColor = LightenColor(Background.Color, lightenPercentage);
            Background = new SolidColorBrush(lightenedColor);
        }

        private static System.Windows.Media.Color LightenColor(System.Windows.Media.Color color, double lightenPercentage)
        {
            // Calculate the lighten amount for each channel (R, G, B)
            byte r = (byte)(color.R + (255 - color.R) * lightenPercentage);
            byte g = (byte)(color.G + (255 - color.G) * lightenPercentage);
            byte b = (byte)(color.B + (255 - color.B) * lightenPercentage);

            return System.Windows.Media.Color.FromRgb(r, g, b);
        }

        /// <summary>
        /// Реализация INotifyPropertyChanged
        /// </summary>
        private void OnPropertyChanged(string propName)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propName));
        }
    }
}
