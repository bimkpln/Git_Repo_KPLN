using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    public sealed class ReportComment : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ReportComment(string message, int type)
        {
            UserName = KPLN_Loader.Preferences.User.SystemName;
            Time = DateTime.Now.ToString();
            Message = message;
            Type = type;
            if (Type == 0 && UserName == KPLN_Loader.Preferences.User.SystemName)
            { VisibleIfUserComment = System.Windows.Visibility.Visible; }
            else
            { VisibleIfUserComment = System.Windows.Visibility.Collapsed; }
        }

        public ReportComment(string user, string date, string message, int type)
        {
            UserName = user;
            Time = date;
            Message = message;
            Type = type;
            if (Type == 0 && UserName == KPLN_Loader.Preferences.User.SystemName)
            { VisibleIfUserComment = System.Windows.Visibility.Visible; }
            else
            { VisibleIfUserComment = System.Windows.Visibility.Collapsed; }
        }

        public static ObservableCollection<ReportComment> ParseComments(string value, ReportInstance instance)
        {
            ObservableCollection<ReportComment> comments = new ObservableCollection<ReportComment>();
            foreach (string comment_data in value.Split(new string[] { Collections.separator_element }, StringSplitOptions.RemoveEmptyEntries))
            {
                List<string> parts = comment_data.Split(new string[] { Collections.separator_sub_element }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (parts.Count != 4) { continue; }
                comments.Add(new ReportComment(parts[0], parts[1], parts[2], int.Parse(parts[3])) { Parent = instance });
            }
            return comments;
        }

        public ReportInstance Parent { get; set; }

        /// <summary>
        /// Имя фамилия из общей БД KPLN в формате "Фамилия Имя"
        /// </summary>
        public string User
        {
            get
            {
                foreach (SQLUserInfo user in KPLN_Loader.Preferences.Users)
                {
                    if (user.SystemName == UserName)
                    {
                        return string.Format("{0} {1}", user.Family, user.Name);
                    }
                }
                return UserName;
            }
        }
        
        public System.Windows.Visibility VisibleIfUserComment { get; set; }
        
        public string UserName { get; set; }
        
        public string Time { get; set; }
        
        public int Type { get; set; }
        
        public string Message { get; set; }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            return string.Join(Collections.separator_sub_element, new string[] { UserName, Time, Message, Type.ToString() });
        }
    }
}
