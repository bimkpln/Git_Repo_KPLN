using Autodesk.Revit.DB;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.WPFItems
{
    /// <summary>
    /// Спец. класс-обертка, для передачи в WPFReport для генерации окна-отчёта
    /// </summary>
    public sealed class WPFEntity : INotifyPropertyChanged
    {
        public WPFEntity(string name)
        {
            Name = name;
        }

        public WPFEntity(Element element)
        {
            Element = element;
            CategoryName = element.Category.Name;
        }

        public WPFEntity(Element element, Status status, string header, string description, bool isZoomElement, bool isApproveElement, string approveComment = null, string info = null) : this(element)
        {
            ElementName = element.Name;
            CurrentStatus = status;
            Header = header;
            Description = description;
            IsZoomElement = isZoomElement;
            IsApproveElement = isApproveElement;
            ApproveComment = approveComment;
            Info = info;

            switch (CurrentStatus)
            {
                case Status.AllmostOk:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 240, 240, 135));
                    Header = "Почти хорошо: " + Header;
                    break;
                case Status.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 220, 90));
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
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 123, 104, 238));
                    Header = "Допустимое: " + Header;
                    break;
            }
        }

        /// <summary>
        /// Создаёт рамку для зумирования
        /// </summary>
        /// <param name="box"></param>
        public void PrepareZoomGeometryExtension(BoundingBoxXYZ box)
        {
            Box = box;
            Centroid = new XYZ((box.Min.X + box.Max.X) / 2, (box.Min.Y + box.Max.Y) / 2, (box.Min.Z + box.Max.Z) / 2);
        }

        public Element Element { get; }

        public string SearchIcon { get; } = "🔍";

        public string ApproveIcon { get; } = "🔓";

        /// <summary>
        /// Имя элемента
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Заголовок элемента
        /// </summary>
        public string Header { get; set; }

        public string CategoryName { get; }

        public string ElementName { get; }

        /// <summary>
        /// Описание элемента
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// Дополнительная инструкция
        /// </summary>
        public string Info { get; }

        private string _approveComment;
        /// <summary>
        /// Комментарий, указанный при подтверждении
        /// </summary>
        public string ApproveComment 
        { 
            get { return _approveComment; }
            set 
            { 
                _approveComment = value;
                OnPropertyChanged("ApproveComment");
            }
        }

        /// <summary>
        /// Текущий статус ошибки
        /// </summary>
        public Status CurrentStatus { get; set; }

        public System.Windows.Visibility Visibility { get; set; }

        /// <summary>
        /// Использовать кастомный зум?
        /// </summary>
        public bool IsZoomElement { get; }

        /// <summary>
        /// Есть возможность подтверждать ошибку?
        /// </summary>
        public bool IsApproveElement { get; }

        private SolidColorBrush _background;
        public SolidColorBrush Background 
        { 
            get { return _background; }
            private set
            {
                _background = value;
                OnPropertyChanged("Background");
            }
        }

        public BoundingBoxXYZ Box { get; private set; }

        public XYZ Centroid { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;
        
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
