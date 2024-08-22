using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.CheckCommandCollections;

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
        /// <summary>
        /// Заголовок элемента
        /// </summary>
        private string _header;

        public WPFEntity(Element element, CheckStatus status, string header, string description, bool isZoomElement, bool isApproveElement, string info = null, string approveComment = null)
        {
            Element = element;
            ElementId = element.Id;
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
            
            CurrentStatus = status;
            ErrorHeader = header;
            _header = header;
            Description = description;
            _approveComment = approveComment;
            Info = info;
            
            IsZoomElement = isZoomElement;
            IsApproveElement = isApproveElement;

            UpdateMainFieldByStatus(status);
        }

        public WPFEntity(IEnumerable<Element> elements, CheckStatus status, string header, string description, bool isZoomElement, bool isApproveElement, string info = null, string approveComment = null)
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

            CurrentStatus = status;
            ErrorHeader = header;
            _header = header;
            Description = description;
            _approveComment = approveComment;
            Info = info;

            IsZoomElement = isZoomElement;
            IsApproveElement = isApproveElement;

            UpdateMainFieldByStatus(status);
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
        /// Id Revit-элемента
        /// </summary>
        public ElementId ElementId { get; }

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
        /// Пользовательский заголовок ошибки (для генерации Header - элемента)
        /// </summary>
        public string ErrorHeader { get; }

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
        public CheckStatus CurrentStatus { get; private set; }

        /// <summary>
        /// Использовать кастомный зум?
        /// </summary>
        public bool IsZoomElement { get; }

        /// <summary>
        /// Есть возможность подтверждать ошибку?
        /// </summary>
        public bool IsApproveElement { get; }

        /// <summary>
        /// Фон элемента
        /// </summary>
        public SolidColorBrush Background
        {
            get => _background;
            private set
            {
                if (value != _background)
                {
                    _background = value;
                    OnPropertyChanged(nameof(Background));
                }
            }
        }

        public BoundingBoxXYZ Box { get; private set; }

        public XYZ Centroid { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Обновление основных визуальных разделителей единицы отчета
        /// </summary>
        /// <param name="status"></param>
        public void UpdateMainFieldByStatus(CheckStatus status)
        {
            CurrentStatus = status;
            switch (status)
            {
                case CheckStatus.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 240, 90));
                    Header = "Обрати внимание: " + Header;
                    break;
                case CheckStatus.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    Header = "Предупреждение: " + Header;
                    break;
                case CheckStatus.Error:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    Header = "Ошибка: " + Header;
                    break;
                case CheckStatus.Approve:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 125, 105, 240));
                    Header = "Допустимое: " + Header;
                    break;
            }
        }

        /// <summary>
        /// Создаёт рамку для зумирования
        /// </summary>
        public void PrepareZoomGeometryExtension(BoundingBoxXYZ box)
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
