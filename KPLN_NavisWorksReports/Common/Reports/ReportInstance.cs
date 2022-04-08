using KPLN_NavisWorksReports.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using static KPLN_NavisWorksReports.Common.Collections;
using static KPLN_Loader.Output.Output;
using System.Data.SQLite;
using Autodesk.Revit.DB;

namespace KPLN_NavisWorksReports.Common.Reports
{
    public class ReportInstance : INotifyPropertyChanged
    {
        private System.Windows.Visibility _IsControllsVisible { get; set; } = System.Windows.Visibility.Visible;
        private bool _IsControllsEnabled { get; set; } = true;
        public ObservableCollection<ReportInstance> SubElements { get; set; } = new ObservableCollection<ReportInstance>();
        public ObservableCollection<ReportComment> _Comments { get; set; } = new ObservableCollection<ReportComment>();
        private string _Path { get; }
        private int _Id { get; set; }
        public int GroupId { get; set; }
        private string _Name { get; set; }
        private int _Element_1_Id { get; set; }
        private int _Element_2_Id { get; set; }
        private string _Element_1_Info { get; set; }
        private string _Element_2_Info { get; set; }
        private string _Point { get; set; }
        private ImageSource _ImageSource { get; set; }
        private byte[] _ImageData { get; set; }
        private byte[] _ImageData_Preview { get; set; }
        private Status _Status { get; set; }
        private SolidColorBrush _Fill { get; set; } = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 255, 255));
        public int Id
        {
            get
            {
                return _Id;
            }
            set
            {
                _Id = value;
                NotifyPropertyChanged();
            }
        }
        public System.Windows.Visibility IsGroup
        {
            get
            {
                if (SubElements.Count != 0)
                { 
                    return  System.Windows.Visibility.Visible;
                }
                return System.Windows.Visibility.Collapsed;
            }
        }
        public bool IsControllsEnabled
        {
            get
            {
                return _IsControllsEnabled;
            }
            set
            {
                _IsControllsEnabled = value;
                NotifyPropertyChanged();
            }
        }
        public System.Windows.Visibility IsControllsVisible
        {
            get
            {
                return _IsControllsVisible;
            }
            set
            {
                _IsControllsVisible = value;
                NotifyPropertyChanged();
            }
        }
        public string Name
        {
            get
            {
                return _Name;
            }
            set
            {
                _Name = value;
                NotifyPropertyChanged();
            }
        }
        public int Element_1_Id
        {
            get
            {
                return _Element_1_Id;
            }
            set
            {
                _Element_1_Id = value;
                NotifyPropertyChanged();
            }
        }
        public int Element_2_Id
        {
            get
            {
                return _Element_2_Id;
            }
            set
            {
                _Element_2_Id = value;
                NotifyPropertyChanged();
            }
        }
        public string Element_1_Info
        {
            get
            {
                return _Element_1_Info;
            }
            set
            {
                _Element_1_Info = value;
                NotifyPropertyChanged();
            }
        }
        public string Element_2_Info
        {
            get
            {
                return _Element_2_Info;
            }
            set
            {
                _Element_2_Info = value;
                NotifyPropertyChanged();
            }
        }
        public string Point
        {
            get
            {
                return _Point;
            }
            set
            {
                _Point = value;
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
            get
            {
                return _ImageSource;
            }
            set
            {
                _ImageSource = value;
                NotifyPropertyChanged();
            }
        }
        public byte[] ImageData
        {
            get
            {
                return _ImageData;
            }
            set
            {
                _ImageData = value;
                NotifyPropertyChanged();
            }
        }
        public byte[] ImageData_Preview
        {
            get
            {
                return _ImageData_Preview;
            }
            set
            {
                _ImageData_Preview = value;
                NotifyPropertyChanged();
            }
        }
        public Status Status
        {
            get
            {
                return _Status;
            }
            set
            {
                if (value == Status.Closed)
                {
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 190, 104));
                }
                if (value == Status.Approved)
                {
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 78, 97, 112));
                }
                if (value == Status.Opened)
                {
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 84, 42));
                }
                _Status = value;
                NotifyPropertyChanged();
            }
        }
        public SolidColorBrush Fill
        {
            get
            {
                return _Fill;
            }
            set
            {
                _Fill = value;
                NotifyPropertyChanged();
            }
        }
        private Stream StreamFromBitmapSource(BitmapSource writeBmp)
        {
            Stream bmp;
            using (bmp = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(writeBmp));
                enc.Save(bmp);
            }

            return bmp;
        }
        public byte[] GetBytes()
        {
            SQLiteConnection db = new SQLiteConnection(string.Format(@"Data Source={0};Version=3;", _Path));
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
        private Stream _ImageStream { get; set; }
        private BitmapImage _BitmapImage { get; set; }
        public void LoadImage()
        {
            ImageSource = null;
            byte[] image_buffer = GetBytes();
            _ImageStream = new MemoryStream(image_buffer, 0, image_buffer.Length);
            _BitmapImage = new BitmapImage();
            _BitmapImage.BeginInit();
            _BitmapImage.StreamSource = _ImageStream;
            _BitmapImage.CacheOption = BitmapCacheOption.OnDemand;
            _BitmapImage.EndInit();
            ImageSource = _BitmapImage;
        }
        public void UnLoadImage()
        {
            ImageSource = null;
            _ImageStream.Dispose();
            _ImageStream = null;
            _BitmapImage = null;
        }
        public ReportInstance(int id, string name, string element_1_id, string element_2_id, string element_1_info, string element_2_info, string point, Status status, string path, int groupId, bool loadImage = true)
        {
            _Path = path;
            Id = id;
            Name = name;
            GroupId = groupId;
            Element_1_Id = int.Parse(element_1_id, System.Globalization.NumberStyles.Integer);
            Element_2_Id = int.Parse(element_2_id, System.Globalization.NumberStyles.Integer);
            Element_1_Info = element_1_info;
            Element_2_Info = element_2_info;
            Point = point;
            Status = status;
            if (loadImage && status == Status.Opened)
            {
                LoadImage();
            }
            Comments = DbController.GetComments(new FileInfo(_Path), this);
        }
        public ReportInstance(int id, string name, string element_1_id, string element_2_id, string element_1_info, string element_2_info, string image, string point, Status status, int groupId, ObservableCollection<ReportComment> comments)
        {
            Id = id;
            Name = name;
            GroupId = groupId;
            Comments = comments;
            try
            {
                Element_1_Id = int.Parse(element_1_id, System.Globalization.NumberStyles.Integer);
            }
            catch (Exception)
            {
                Element_1_Id = -1;
            }
            try
            {
                Element_2_Id = int.Parse(element_2_id, System.Globalization.NumberStyles.Integer);
            }
            catch (Exception)
            {
                Element_2_Id = -1;
            }
            Element_1_Info = element_1_info;
            Element_2_Info = element_2_info;
            using (Stream image_stream = System.IO.File.Open(image, FileMode.Open))
            {
                ImageData = SystemTools.ReadFully(image_stream);
            }
            using (Stream image_stream = System.IO.File.Open(image, FileMode.Open))
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
                                    //string comments = rdr.GetString(8);
                                    int groupId = rdr.GetInt32(9);
                                    reports.Add(new ReportInstance(id, name, el1_id, el2_id, el1_name, el2_name, point, status, path, groupId, (Num < 200 && groupId == -1)));
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

        public ObservableCollection<ReportComment> Comments
        {
            get
            {
                return _Comments;
            }
            set
            {
                _Comments = value;
                NotifyPropertyChanged();
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public void AddComment(string message, int type)
        {
            try
            {
                DbController.AddComment(message, new FileInfo(_Path), this, type);
                Comments = DbController.GetComments(new FileInfo(_Path), this);
            }
            catch (Exception)
            { }
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
        public void RemoveComment(ReportComment comment)
        {
            try
            {
                DbController.RemoveComment(comment, new FileInfo(_Path), this);
                Comments = DbController.GetComments(new FileInfo(_Path), this);
            }
            catch (Exception)
            { }
        }
    }
}
