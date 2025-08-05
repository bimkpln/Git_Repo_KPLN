using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using static KPLN_Clashes_Ribbon.Core.ClashesMainCollection;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    public sealed class ReportItemComment
    {
        public ReportItemComment(string message)
        {
            UserSystemName = DBMainService.CurrentDBUser.SystemName;
            Time = DateTime.Now.ToString();
            Message = message;
        }

        public ReportItemComment(string user, string date, string message)
        {
            UserSystemName = user;
            Time = date;
            Message = message;
        }

        public static ObservableCollection<ReportItemComment> ParseComments(string value, ReportItem instance)
        {
            ObservableCollection<ReportItemComment> comments = new ObservableCollection<ReportItemComment>();
            foreach (string comment_data in value.Split(new string[] { ClashesMainCollection.StringSeparatorItem }, StringSplitOptions.RemoveEmptyEntries))
            {
                List<string> parts = comment_data.Split(new string[] { ClashesMainCollection.StringSeparatorSubItem }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (parts.Count != 3) { continue; }
                comments.Add(new ReportItemComment(parts[0], parts[1], parts[2]) { Parent = instance });
            }
            return comments;
        }

        public ReportItem Parent { get; set; }

        /// <summary>
        /// Фамилия имя из общей БД KPLN в формате "Фамилия Имя"
        /// </summary>
        public string UserFullName
        {
            get
            {
                DBUser dBUser = DBMainService.UserDbService.GetDBUser_ByUserName(UserSystemName);
                return $"{dBUser.Surname} {dBUser.Name}";
            } 
        }

        /// <summary>
        /// Имя учетной записи Windows (из общей БД KPLN)
        /// </summary>
        public string UserSystemName { get; private set; }

        public string Time { get; set; }

        public string Message { get; set; }

        public override string ToString()
        {
            return string.Join(ClashesMainCollection.StringSeparatorSubItem, new string[] { UserSystemName, Time, Message });
        }
    }
}
