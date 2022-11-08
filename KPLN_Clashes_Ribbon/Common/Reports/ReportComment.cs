using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Common.Reports
{
    public sealed class ReportComment : INotifyPropertyChanged
    {
        public ReportInstance Parent { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public string User
        {
            get
            {
                foreach (SQLUserInfo user in KPLN_Loader.Preferences.Users)
                {
                    if (user.SystemName == UserSystemName)
                    {
                        if (user.Surname != string.Empty)
                        {
                            return string.Format("{0} {1}.{2}.", user.Family, user.Name[0], user.Surname[0]);
                        }
                        else
                        {
                            return string.Format("{0} {1}", user.Family, user.Name);
                        }
                    }
                }
                return UserSystemName;
            }
        }
        public System.Windows.Visibility VisibleIfUserComment { get; set; }
        public string UserSystemName { get; set; }
        public string Time { get; set; }
        public int Type { get; set; }
        public string Message { get; set; }
        public ReportComment(string message, int type)
        {
            UserSystemName = KPLN_Loader.Preferences.User.SystemName;
            Time = DateTime.Now.ToString();
            Message = message;
            Type = type;
            if (Type == 0 && UserSystemName == KPLN_Loader.Preferences.User.SystemName)
            { VisibleIfUserComment = System.Windows.Visibility.Visible; }
            else
            { VisibleIfUserComment = System.Windows.Visibility.Collapsed; }
        }
        public ReportComment(string user, string date, string message, int type)
        {
            UserSystemName = user;
            Time = date;
            Message = message;
            Type = type;
            if (Type == 0 && UserSystemName == KPLN_Loader.Preferences.User.SystemName)
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
        public override string ToString()
        {
            return string.Join(Collections.separator_sub_element, new string[] { UserSystemName, Time, Message, Type.ToString() });
        }
    }
}
