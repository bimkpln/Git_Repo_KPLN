﻿using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    /// <summary>
    /// Данные по каждому отчету в отдельности из таблиц отдельных отчетов
    /// </summary>
    public sealed class ReportItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Вручную прописал отделы из БД, т.к. они не меняются
        /// </summary>
        private ObservableCollection<SubDepartmentBtn> _subDepartmentBtns = new ObservableCollection<SubDepartmentBtn>()
        {
            new SubDepartmentBtn(2, "АР"),
            new SubDepartmentBtn(3, "КР"),
            new SubDepartmentBtn(4, "ОВ"),
            new SubDepartmentBtn(5, "ВК"),
            new SubDepartmentBtn(6, "ЭОМ"),
            new SubDepartmentBtn(7, "СС"),
            new SubDepartmentBtn(20, "ИТП"),
            new SubDepartmentBtn(21, "ПТ"),
            new SubDepartmentBtn(22, "АВ", "Подраздел СС: Автоматизация"),
            new SubDepartmentBtn(99, "✖", "Сбросить делегирование и вернуть статус пересечения «Открытое»"),
        };

        private int _statusId;
        private string _comments;
        private KPItemStatus _status;
        private int _delegatedDepartmentId;
        private System.Windows.Visibility _isControllsVisible = System.Windows.Visibility.Visible;
        private bool _isControllsEnabled = true;
        private SolidColorBrush _fill = new SolidColorBrush(Color.FromArgb(255, 255, 255, 255));
        private ImageSource _imageSource;
        private Stream _imageStream;
        private BitmapImage _bitmapImage;
        private ObservableCollection<ReportItemComment> _сommentCollection = new ObservableCollection<ReportItemComment>();

        /// <summary>
        /// Конструктор для Dapper (он по-умолчанию использует его, когда мапит данные из БД)
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
            int repId,
            string name,
            string element_1_id,
            string element_2_id,
            string element_1_info,
            string element_2_info,
            string image,
            string point,
            KPItemStatus status,
            string reportParentGroupComments,
            string reportItemComments,
            int parentGroupId)
        {
            Id = id;
            ReportGroupId = repGroupId;
            ReportId = repId;
            Name = name;
            ParentGroupId = parentGroupId;

            if (int.TryParse(element_1_id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id_1))
                Element_1_Id = id_1;
            else
                Element_1_Id = -1;

            if (int.TryParse(element_2_id, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id_2))
                Element_2_Id = id_2;
            else
                Element_2_Id = -1;
            
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
            ReportParentGroupComments = reportParentGroupComments;
            ReportItemComments = reportItemComments;
        }

        #region Данные из БД
        [Key]
        public int Id { get; set; }

        public int ReportGroupId { get; set; }

        public int ReportId { get; set; }

        public string Name { get; set; }

        public byte[] Image { get; set; }

        public int Element_1_Id { get; set; }

        public int Element_2_Id { get; set; }

        public string Element_1_Info { get; set; }

        public string Element_2_Info { get; set; }

        public string Point { get; set; }

        /// <summary>
        /// ID статуса. Нужен для маппинга внутри Dapper. Set влияет на параметр Status
        /// </summary>
        public int StatusId
        {
            get
            {
                _statusId = (int)Status;
                return _statusId;
            }
            set
            {
                _statusId = value;
                Status = (KPItemStatus)_statusId;
            }
        }

        public string ReportParentGroupComments { get; set; }
        
        public System.Windows.Visibility ReportParentGroupCommentsVisibility 
        {
            get
            {
                if (string.IsNullOrEmpty(ReportParentGroupComments))
                    return System.Windows.Visibility.Collapsed;
                
                return System.Windows.Visibility.Visible;
            }
        }

        public string ReportItemComments { get; set; }

        public System.Windows.Visibility ReportItemCommentsVisibility
        {
            get
            {
                if (string.IsNullOrEmpty(ReportItemComments))
                    return System.Windows.Visibility.Collapsed;

                return System.Windows.Visibility.Visible;
            }
        }

        /// <summary>
        /// Если коллизия в группе - ссылка на id данной группы, иначе значение -1 (приходит из настроек БД)
        /// </summary>
        public int ParentGroupId { get; set; }

        public int DelegatedDepartmentId
        {
            get => _delegatedDepartmentId;
            set 
            { 
                _delegatedDepartmentId = value;
                NotifyPropertyChanged();

                //Вёрстка на лету цвета кнопки делегации
                SubDepartmentBtn subDep = _subDepartmentBtns.FirstOrDefault(sdb => sdb.Id == _delegatedDepartmentId);
                if (subDep != null)
                    subDep.DelegateBtnBackground = Brushes.Aqua;
            }
        }


        /// <summary>
        /// Закодированный коммент из БД. Нужен для маппинга внутри Dapper. Set влияет на отображение в окне (ReportItemCommentCollection)
        /// </summary>
        public string Comments
        {
            get => _comments;
            set
            {
                _comments = value;
                CommentCollection = ReportItemComment.ParseComments(_comments, this);
            }
        }
        #endregion

        #region Дополнительная визуализация
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

        public System.Windows.Visibility IsGroup
        {
            get
            {
                if (SubElements.Count != 0)
                    return System.Windows.Visibility.Visible;
                
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

        /// <summary>
        /// Визуализация комментариев в окне
        /// </summary>
        public ObservableCollection<ReportItemComment> CommentCollection
        {
            get => _сommentCollection;
            set
            {
                _сommentCollection = value;
                NotifyPropertyChanged();
            }
        }

        /// <summary>
        /// Коллекция кастомных отделов КПЛН (только внутри данного плагина)
        /// </summary>
        public ObservableCollection<SubDepartmentBtn> SubDepartmentBtns 
        {
            get => _subDepartmentBtns;
            set
            {
                _subDepartmentBtns = value;
                NotifyPropertyChanged();
            }
        } 

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
    }
}
