using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static KPLN_Clashes_Ribbon.Common.Collections;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    /// <summary>
    /// Данные по каждому отчету в отдельности из таблиц отдельных отчетов
    /// </summary>
    public sealed class ReportInstance : INotifyPropertyChanged
    {
        public ObservableCollection<ReportComment> _comments = new ObservableCollection<ReportComment>();

        public ObservableCollection<ReportInstance> _subElements = new ObservableCollection<ReportInstance>();

        public ObservableCollection<SubDepartmentBtn> _subDepartmentBtns = new ObservableCollection<SubDepartmentBtn>()
        {
            new SubDepartmentBtn(1, "АР", "Разделы АР"),
            new SubDepartmentBtn(2, "КР", "Разделы КР"),
            new SubDepartmentBtn(3, "ВК", "Разделы АУПТ, ВК, НС"),
            new SubDepartmentBtn(4, "ОВ", "Разделы ИТП, ОВиК"),
            new SubDepartmentBtn(5, "СС", "Разделы СС"),
            new SubDepartmentBtn(6, "ЭОМ", "Разделы ЭОМ"),
            new SubDepartmentBtn(7, "✖", "Сбросить делегирование и вернуть статус пересечения «Открытое»"),
        };

        public event PropertyChangedEventHandler PropertyChanged;

        private readonly string _path;

        private int _id;

        private string _name;

        private byte[] _imageData;

        private byte[] _imageData_Preview;

        private ImageSource _imageSource;

        private int _element_1_Id;

        private int _element_2_Id;

        private string _element_1_Info;

        private string _element_2_Info;

        private string _point;

        private Status _status;

        private int _delegatedDepartmentId;

        private System.Windows.Visibility _isControllsVisible = System.Windows.Visibility.Visible;

        private bool _isControllsEnabled = true;

        private SolidColorBrush _fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255, 255));
        
        private Stream _imageStream;

        private BitmapImage _bitmapImage;

        /// <summary>
        /// Конструктор для генерации отчетов из БД
        /// </summary>
        public ReportInstance(
            int id,
            string name,
            string element_1_id,
            string element_2_id,
            string element_1_info,
            string element_2_info,
            string point,
            Status status,
            string path,
            int groupId,
            int delegatedDepartmentId,
            bool loadImage)
        {
            _path = path;

            Id = id;
            Name = name;
            GroupId = groupId;
            Element_1_Id = int.Parse(element_1_id, System.Globalization.NumberStyles.Integer);
            Element_2_Id = int.Parse(element_2_id, System.Globalization.NumberStyles.Integer);
            Element_1_Info = element_1_info;
            Element_2_Info = element_2_info;
            Point = point;
            Status = status;
            Comments = DbController.GetComments(new FileInfo(_path), this);
            DelegatedDepartmentId = delegatedDepartmentId;

            if (loadImage && status == Status.Opened)
            { LoadImage(); }

            // Генерация кнопок делегирования
            foreach (SubDepartmentBtn sdBtn in _subDepartmentBtns)
            {
                if (sdBtn.Id == DelegatedDepartmentId)
                { sdBtn.SetBinding(this, Brushes.Aqua); }
                else
                { sdBtn.SetBinding(this, Brushes.Transparent); }
            }
            
        }

        /// <summary>
        /// Конструктор для генерации отчетов из html-отчетов
        /// </summary>
        public ReportInstance(
            int id,
            string name,
            string element_1_id,
            string element_2_id,
            string element_1_info,
            string element_2_info,
            string image,
            string point,
            Status status,
            int groupId,
            ObservableCollection<ReportComment> comments)
        {
            Id = id;
            Name = name;
            GroupId = groupId;
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
                ImageData = SystemTools.ReadFully(image_stream);
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

        public int Id
        {
            get { return _id; }
            set
            {
                _id = value;
                NotifyPropertyChanged();
            }
        }

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
            get { return _isControllsEnabled; }
            set
            {
                _isControllsEnabled = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility IsControllsVisible
        {
            get
            {
                return _isControllsVisible;
            }
            set
            {
                _isControllsVisible = value;
                NotifyPropertyChanged();
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                NotifyPropertyChanged();
            }
        }

        public int Element_1_Id
        {
            get { return _element_1_Id; }
            set
            {
                _element_1_Id = value;
                NotifyPropertyChanged();
            }
        }

        public int Element_2_Id
        {
            get { return _element_2_Id; }
            set
            {
                _element_2_Id = value;
                NotifyPropertyChanged();
            }
        }

        public string Element_1_Info
        {
            get { return _element_1_Info; }
            set
            {
                _element_1_Info = value;
                NotifyPropertyChanged();
            }
        }

        public string Element_2_Info
        {
            get { return _element_2_Info; }
            set
            {
                _element_2_Info = value;
                NotifyPropertyChanged();
            }
        }

        public string Point
        {
            get { return _point; }
            set
            {
                _point = value;
                NotifyPropertyChanged();
            }
        }

        public System.Windows.Visibility PlacePointVisibility
        {
            get
            {
                if (Point == "NONE")
                { return System.Windows.Visibility.Collapsed; }
                return System.Windows.Visibility.Visible;
            }
        }

        public ImageSource ImageSource
        {
            get { return _imageSource; }
            set
            {
                _imageSource = value;
                NotifyPropertyChanged();
            }
        }

        public byte[] ImageData
        {
            get { return _imageData; }
            set
            {
                _imageData = value;
                NotifyPropertyChanged();
            }
        }

        public byte[] ImageData_Preview
        {
            get { return _imageData_Preview; }
            set
            {
                _imageData_Preview = value;
                NotifyPropertyChanged();
            }
        }

        public int GroupId { get; set; }

        public Status Status
        {
            get { return _status; }
            set
            {
                switch (value)
                {
                    case Status.Closed:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 0, 190, 104));
                        break;
                    case Status.Approved:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 78, 97, 112));
                        break;
                    case Status.Delegated:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 75, 0, 130));
                        break;
                    case Status.Opened:
                        Fill = new SolidColorBrush(Color.FromArgb(255, 255, 84, 42));
                        break;
                }

                _status = value;

                NotifyPropertyChanged();
            }
        }

        public int DelegatedDepartmentId
        {
            get { return _delegatedDepartmentId; }
            private set { _delegatedDepartmentId = value; }
        }

        public ObservableCollection<ReportComment> Comments
        {
            get { return _comments; }
            set
            {
                _comments = value;
                NotifyPropertyChanged();
            }
        }

        public ObservableCollection<SubDepartmentBtn> SubDepartmentBtns
        {
            get { return _subDepartmentBtns; }
            private set { _subDepartmentBtns = value; }
        }

        public ObservableCollection<ReportInstance> SubElements
        {
            get { return _subElements; }
            set { _subElements = value; }
        }

        public SolidColorBrush Fill
        {
            get { return _fill; }
            set
            {
                _fill = value;
                NotifyPropertyChanged();
            }
        }

        public static ObservableCollection<ReportInstance> GetReportInstances(string path)
        {
            ObservableCollection<ReportInstance> reports = new ObservableCollection<ReportInstance>();
            try
            {
                SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", path));
                try
                {
                    db.Open();
                    int Num = 0;
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT ID FROM Reports", db))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                Num++;
                            }
                        }
                    }
                    using (SQLiteCommand cmd = new SQLiteCommand("SELECT * FROM Reports", db))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                if (rdr.GetInt32(0) == -1) { continue; }
                                try
                                {
                                    int id = rdr.GetInt32(0);
                                    string name = rdr.GetString(1);
                                    string el1_name = rdr.GetString(4).Split('|').Last();
                                    string el2_name = rdr.GetString(5).Split('|').Last();
                                    string el1_id = rdr.GetString(4).Split('|').First();
                                    string el2_id = rdr.GetString(5).Split('|').First();
                                    string point = rdr.GetString(6);
                                    Status status = Status.Opened;
                                    int status_int = rdr.GetInt32(7);
                                    if (status_int == 0)
                                    { status = Status.Closed; }
                                    if (status_int == 1)
                                    { status = Status.Approved; }
                                    if (status_int == 2)
                                    { status = Status.Delegated; }
                                    int groupId = rdr.GetInt32(9);

                                    // Исключение необходимо для старых отчетов, до добавления делегирования
                                    int departentId = -1;
                                    try
                                    { departentId = rdr.GetInt32(10); }
                                    catch (InvalidCastException) { }
                                    catch (IndexOutOfRangeException) { }

                                    reports.Add(new ReportInstance(
                                        id,
                                        name,
                                        el1_id,
                                        el2_id,
                                        el1_name,
                                        el2_name,
                                        point,
                                        status,
                                        path,
                                        groupId,
                                        departentId,
                                        (Num < 200 && groupId == -1)));
                                }
                                catch (Exception e)
                                {
                                    PrintError(e);
                                }

                            }
                        }
                    }
                    db.Close();
                }
                catch (Exception)
                {
                    db.Close();
                }
            }
            catch (Exception) { }
            ObservableCollection<ReportInstance> result_reports = new ObservableCollection<ReportInstance>();
            foreach (ReportInstance i in reports)
            {
                if (i.GroupId == -1)
                {
                    result_reports.Add(i);
                }
            }
            foreach (ReportInstance i in reports)
            {
                if (i.GroupId != -1)
                {
                    foreach (ReportInstance z in result_reports)
                    {
                        if (i.GroupId == z.Id)
                        {
                            z.SubElements.Add(i);
                        }
                    }
                }
            }
            return result_reports;
        }

        public static string GetCommentsString(ObservableCollection<ReportComment> comments)
        {
            List<string> value_parts = new List<string>();
            foreach (ReportComment comment in comments)
            {
                value_parts.Add(comment.ToString());
            }
            string value = string.Join(Collections.separator_element, value_parts);
            return value;
        }

        public void LoadImage()
        {
            ImageSource = null;
            byte[] image_buffer = GetBytes();
            _imageStream = new MemoryStream(image_buffer, 0, image_buffer.Length);
            _bitmapImage = new BitmapImage();
            _bitmapImage.BeginInit();
            _bitmapImage.StreamSource = _imageStream;
            _bitmapImage.CacheOption = BitmapCacheOption.OnDemand;
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

        public void RemoveComment(ReportComment comment)
        {
            try
            {
                DbController.RemoveComment(comment, new FileInfo(_path), this);
                Comments = DbController.GetComments(new FileInfo(_path), this);
            }
            catch (Exception)
            { }
        }

        public byte[] GetBytes()
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", _path));
            db.Open();
            using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT IMAGE FROM Reports WHERE ID={0}", Id), db))
            {
                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        try
                        {
                            byte[] buffer = new byte[512 * 1024];
                            rdr.GetBytes(0, 0, buffer, 0, buffer.Length);
                            return buffer;
                        }
                        catch (Exception e)
                        {
                            PrintError(e);
                        }

                    }
                }
            }
            db.Close();
            return null;
        }
    }
}
