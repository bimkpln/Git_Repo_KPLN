using Autodesk.Revit.DB;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    public class WPFDisplayItem : INotifyPropertyChanged
    {
        public BoundingBoxXYZ Box { get; private set; }

        public XYZ Centroid { get; private set; }

        public int CategoryId { get; }

        public string Icon { get; set; }

        public Element Element { get; set; }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }

        public int ElementId
        {
            get { return _elementId; }
            set
            {
                _elementId = value;
                NotifyPropertyChanged();
            }
        }

        public string Category
        {
            get { return _category; }
            set
            {
                _category = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility Visibility
        {
            get { return _visibility; }
            set
            {
                _visibility = value;
                NotifyPropertyChanged();
            }
        }

        public SolidColorBrush Background
        {
            get { return _background; }
            set
            {
                _background = value;
                NotifyPropertyChanged();
            }
        }

        public string Header
        {
            get { return _header; }
            set
            {
                _header = value;
                NotifyPropertyChanged();
            }
        }

        public string Description
        {
            get { return _description; }
            set
            {
                _description = value;
                NotifyPropertyChanged();
            }
        }

        public string ToolTip
        {
            get { return _toolTip; }
            set
            {
                _toolTip = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableCollection<WPFDisplayItem> Collection
        {
            get { return _collection; }
            set
            {
                _collection = value;
                NotifyPropertyChanged();
            }
        }

        public bool IsEnabled
        {
            get { return _isEnabled; }
            set
            {
                _isEnabled = value;
                NotifyPropertyChanged();
            }
        }

        private ObservableCollection<WPFDisplayItem> _collection;

        private int _elementId;

        private string _toolTip;

        private string _header;

        private string _description;

        private string _category;

        private string _name;

        private SolidColorBrush _background;

        private bool _isEnabled;

        private System.Windows.Visibility _visibility;

        public WPFDisplayItem(int categoryId, StatusExtended status, string icon = "🔍")
        {
            CategoryId = categoryId;
            switch (status)
            {
                case StatusExtended.LittleWarning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 220, 90));
                    break;
                case StatusExtended.Warning:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 180, 90));
                    break;
                case StatusExtended.Critical:
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 195, 0));
                    break;
            }
            Icon = icon;
        }

        public WPFDisplayItem(int categoryId, StatusExtended status, int elementId) : this(categoryId, status)
        {
            ElementId = elementId;
        }

        public void SetZoomParams(Element element, BoundingBoxXYZ box)
        {
            Element = element;
            if (box != null)
            {
                Box = box;
                Centroid = new XYZ((box.Max.X + box.Min.X) / 2, (box.Max.Y + box.Min.Y) / 2, (box.Max.Z + box.Min.Z) / 2);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
