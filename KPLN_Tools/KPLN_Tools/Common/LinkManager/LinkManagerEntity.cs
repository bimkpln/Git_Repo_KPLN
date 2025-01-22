using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;

namespace KPLN_Tools.Common.LinkManager
{
    public class LinkManagerEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _linkName;
        private string _linkPath;

        private EntityStatus _currentEntStatus;
        private SolidColorBrush _fillColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255, 255));

        public LinkManagerEntity() { }

        public LinkManagerEntity(string name, string path)
        {
            LinkName = name;
            LinkPath = path;

            CurrentEntStatus = EntityStatus.Ok;
        }

        /// <summary>
        /// Имя связи
        /// </summary>
        public string LinkName
        {
            get => _linkName;
            set
            {
                _linkName = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Путь к связи
        /// </summary>
        public string LinkPath
        {
            get => _linkPath;
            set
            {
                _linkPath = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Цвет рамки для окна WPF
        /// </summary>
        public SolidColorBrush FillColor
        {
            get => _fillColor;
            set
            {
                _fillColor = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Статус сущности для окна WPF
        /// </summary>
        public EntityStatus CurrentEntStatus
        {
            get => _currentEntStatus;
            set
            {
                _currentEntStatus = value;
                ResetFillColor();
                NotifyPropertyChanged();
            }
        }

        protected void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ResetFillColor()
        {
            switch (CurrentEntStatus)
            {
                case EntityStatus.Ok:
                    FillColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255, 255));
                    break;
                case EntityStatus.Error:
                    FillColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 165, 0));
                    break;
                case EntityStatus.MarkedAsFinal:
                    FillColor = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 128, 0));
                    break;
            }
        }
    }
}
