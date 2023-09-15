using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    public sealed class ReportItemComment : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public ReportItemComment(string message, int type)
        {
            UserSystemName = KPLN_Loader.Application.CurrentRevitUser.SystemName;
            Time = DateTime.Now.ToString();
            Message = message;
            VisibleIfUserComment = System.Windows.Visibility.Visible;
        }

        public ReportItemComment(string user, string date, string message, int type)
        {
            UserSystemName = user;
            Time = date;
            Message = message;
            VisibleIfUserComment = System.Windows.Visibility.Visible;
        }

        public static ObservableCollection<ReportItemComment> ParseComments(string value, ReportItem instance)
        {
            ObservableCollection<ReportItemComment> comments = new ObservableCollection<ReportItemComment>();
            foreach (string comment_data in value.Split(new string[] { ClashesMainCollection.StringSeparatorItem }, StringSplitOptions.RemoveEmptyEntries))
            {
                List<string> parts = comment_data.Split(new string[] { ClashesMainCollection.StringSeparatorSubItem }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (parts.Count != 4) { continue; }
                comments.Add(new ReportItemComment(parts[0], parts[1], parts[2], int.Parse(parts[3])) { Parent = instance });
            }
            return comments;
        }

        public ReportItem Parent { get; set; }

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
        
        public string Message { get; set; }

        public void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            return string.Join(ClashesMainCollection.StringSeparatorSubItem, new string[] { UserSystemName, Time, Message });
        }
    }
}
