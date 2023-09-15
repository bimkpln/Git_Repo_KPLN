using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    /// <summary>
    /// Данные по каждому отчету в отдельности из таблиц отдельных отчетов
    /// </summary>
    public sealed class ReportItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly string _path;
        private KPItemStatus _status;
        private int _delegatedDepartmentId;
        private System.Windows.Visibility _isControllsVisible = System.Windows.Visibility.Visible;
        private bool _isControllsEnabled = true;
        private SolidColorBrush _fill = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
        private ImageSource _imageSource;
        private Stream _imageStream;
        private BitmapImage _bitmapImage;
        private ObservableCollection<ReportItemComment> _comments = new ObservableCollection<ReportItemComment>();

        /// <summary>
        /// Конструктор-заглушка для Dapper (он по-умолчанию использует его, когда мапит данные из БД)
        /// </summary>
        public ReportItem()
        {
        }

        /// <summary>
        /// Конструктор для генерации отчетов из html-отчетов
        /// </summary>
        public ReportItem(
            int id,
            int repGroupId,
            string name,
            string element_1_id,
            string element_2_id,
            string element_1_info,
            string element_2_info,
            string image,
            string point,
            KPItemStatus status,
            int parentGroupId,
            ObservableCollection<ReportItemComment> comments)
        {
            Id = id;
            ReportGroupId = repGroupId;
            Name = name;
            ParentGroupId = parentGroupId;
            Comments = comments;

            try
            { Element_1_Id = int.Parse(element_1_id, System.Globalization.NumberStyles.Integer); }
            catch (Exception)
            { Element_1_Id = -1; }

            try
            { Element_2_Id = int.Parse(element_2_id, System.Globalization.NumberStyles.Integer); }
            catch (Exception)
            { Element_2_Id = -1; }

            Element_1_Info = element_1_info;
            Element_2_Info = element_2_info;

            using (Stream image_stream = File.Open(image, FileMode.Open))
            {
                Image = SystemTools.ReadFully(image_stream);
            }

            using (Stream image_stream = File.Open(image, FileMode.Open))
            {
                var imageSource = new BitmapImage();
                imageSource.BeginInit();
                imageSource.StreamSource = image_stream;
                imageSource.CacheOption = BitmapCacheOption.Default;
                imageSource.EndInit();
                ImageSource = imageSource;
            }

            Point = point;
            Status = status;
        }

        #region Данные из БД
        [Key]
        public int Id { get; set; }

        public int ReportGroupId { get; set; }
        
        public string Name { get; set; }

        public byte[] Image { get; set; }

        public string Element_1_Info { get; set; }

        public string Element_2_Info { get; set; }

        public string Point { get; set; }

        public KPItemStatus Status
        {
            get => _status;
            set
            {
                switch (value)
                {
                    case KPItemStatus.Closed:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 0, 190, 104));
                        break;
                    case KPItemStatus.Approved:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 78, 97, 112));
                        break;
                    case KPItemStatus.Delegated:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 75, 0, 130));
                        break;
                    case KPItemStatus.Opened:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 255, 84, 42));
                        break;
                }

                _status = value;

                NotifyPropertyChanged();
            }
        }
        
        /// <summary>
        /// Если коллизия в группе - ссылка на id данной группы, иначе значение -1 (приходит из настроек БД)
        /// </summary>
        public int ParentGroupId { get; set; }

        public int DelegatedDepartmentId
        {
            get => _delegatedDepartmentId;
            private set { _delegatedDepartmentId = value; }
        }

        public ObservableCollection<ReportItemComment> Comments
        {
            get => _comments;
            set
            {
                _comments = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        #region Дополнительная визуализация
        public int Element_1_Id { get; set; }

        public int Element_2_Id { get; set; }

        public System.Windows.Visibility IsGroup
        {
            get
            {
                if (SubElements.Count != 0)
                {
                    return System.Windows.Visibility.Visible;
                }
                return System.Windows.Visibility.Collapsed;
            }
        }

        public bool IsControllsEnabled
        {
            get => _isControllsEnabled;
            set
            {
                _isControllsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility IsControllsVisible
        {
            get => _isControllsVisible;
            set
            {
                _isControllsVisible = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility PlacePointVisibility
        {
            get
            {
                if (Point == "NONE")
                    return System.Windows.Visibility.Collapsed; 
                
                return System.Windows.Visibility.Visible;
            }
        }

        public ImageSource ImageSource
        {
            get => _imageSource; 
            set
            {
                if (_imageSource != value)
                {
                    _imageSource = value;
                    NotifyPropertyChanged();
                }
            }
        }

        public ObservableCollection<SubDepartmentBtn> SubDepartmentBtns { get; } = new ObservableCollection<SubDepartmentBtn>()
        {
            new SubDepartmentBtn(1, "АР", "Разделы АР"),
            new SubDepartmentBtn(2, "КР", "Разделы КР"),
            new SubDepartmentBtn(3, "ВК", "Разделы АУПТ, ВК, НС"),
            new SubDepartmentBtn(4, "ОВ", "Разделы ИТП, ОВиК"),
            new SubDepartmentBtn(5, "СС", "Разделы СС"),
            new SubDepartmentBtn(6, "ЭОМ", "Разделы ЭОМ"),
            new SubDepartmentBtn(7, "✖", "Сбросить делегирование и вернуть статус пересечения «Открытое»"),
        };

        /// <summary>
        /// Коллекция субэлементов, если коллизия в группе
        /// </summary>
        public ObservableCollection<ReportItem> SubElements { get; set; } = new ObservableCollection<ReportItem>();

        public SolidColorBrush Fill
        {
            get { return _fill; }
            set
            {
                _fill = value;
                NotifyPropertyChanged();
            }
        }
        #endregion

        public static string GetCommentsString(ObservableCollection<ReportItemComment> comments)
        {
            List<string> value_parts = new List<string>();
            foreach (ReportItemComment comment in comments)
            {
                value_parts.Add(comment.ToString());
            }
            string value = string.Join(ClashesMainCollection.StringSeparatorItem, value_parts);
            return value;
        }

        public void LoadImage(byte[] image_buffer)
        {
            ImageSource = null;
            
            _imageStream = new MemoryStream(image_buffer);
            _bitmapImage = new BitmapImage();
            _bitmapImage.BeginInit();
            _bitmapImage.StreamSource = _imageStream;
            _bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            _bitmapImage.EndInit();
            
            ImageSource = _bitmapImage;
        }

        public void UnLoadImage()
        {
            ImageSource = null;
            _imageStream.Dispose();
            _imageStream = null;
            _bitmapImage = null;
        }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void AddComment(string message, int type)
        {
            try
            {
                DbController.AddComment(message, new FileInfo(_path), this, type);
                Comments = DbController.GetComments(new FileInfo(_path), this);
            }
            catch (Exception)
            { }
        }

        public void RemoveComment(ReportItemComment comment)
        {
            try
            {
                DbController.RemoveComment(comment, new FileInfo(_path), this);
                Comments = DbController.GetComments(new FileInfo(_path), this);
            }
            catch (Exception)
            { }
        }
    }
}
