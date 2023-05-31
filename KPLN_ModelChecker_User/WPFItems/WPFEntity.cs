using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

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

        public WPFEntity(Element element, Status status, string header, string description, bool isZoomElement, bool isApproveElement, string approveComment = null, string info = null)
        {
            Element = element;
            ElementId = element.Id;
            CategoryName = element.Category.Name;
            ElementName = element.Name;
            CurrentStatus = status;
            _header = header;
            Description = description;
            IsZoomElement = isZoomElement;
            IsApproveElement = isApproveElement;
            _approveComment = approveComment;
            Info = info;

            UpdateMainFieldByStatus(status);
        }

        /// <summary>
        /// Revit-элемент
        /// </summary>
        public Element Element { get; }

        /// <summary>
        /// Id Revit-элемента
        /// </summary>
        public ElementId ElementId { get; }

        public string SearchIcon { get; } = "🔍";

        public string ApproveIcon { get; } = "🔓";

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
        public string ElementName { get; }

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
        public Status CurrentStatus { get; private set; }

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
        public void UpdateMainFieldByStatus(Status status)
        {
            CurrentStatus = status;
            switch (status)
            {
                case Status.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 240, 90));
                    Header = "Обрати внимание: " + Header;
                    break;
                case Status.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    Header = "Предупреждение: " + Header;
                    break;
                case Status.Error:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 125, 125));
                    Header = "Ошибка: " + Header;
                    break;
                case Status.Approve:
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
            Box = box;
            Centroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
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
