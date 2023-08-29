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
            UserSystemName = KPLN_Loader.Application.CurrentRevitUser.SystemName;
            Time = DateTime.Now.ToString();
            Message = message;
            Type = type;
            
            if (Type == 0)
                VisibleIfUserComment = System.Windows.Visibility.Visible;
            else
                VisibleIfUserComment = System.Windows.Visibility.Collapsed;
        }

        public ReportComment(string user, string date, string message, int type)
        {
            UserSystemName = user;
            Time = date;
            Message = message;
            Type = type;
            
            if (Type == 0 )
                VisibleIfUserComment = System.Windows.Visibility.Visible; 
            else
                VisibleIfUserComment = System.Windows.Visibility.Collapsed;
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
        /// Фамилия имя из общей БД KPLN в формате "Фамилия Имя"
        /// </summary>
        public string UserFullName
        {
            get => $"{KPLN_Loader.Application.CurrentRevitUser.Surname} {KPLN_Loader.Application.CurrentRevitUser.Name}";
        }
        
        public System.Windows.Visibility VisibleIfUserComment { get; set; }

        /// <summary>
        /// Имя учетной записи Windows ( из общей БД KPLN)
        /// </summary>
        public string UserSystemName { get; private set; }
        
        public string Time { get; set; }
        
        public int Type { get; set; }
        
        public string Message { get; set; }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            return string.Join(Collections.separator_sub_element, new string[] { UserSystemName, Time, Message, Type.ToString() });
        }
    }
}
