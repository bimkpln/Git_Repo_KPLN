using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace KPLN_Clashes_Ribbon.Core.Reports
{
    public sealed class ReportItemComment : INotifyPropertyChanged
    {
        private static readonly ConcurrentDictionary<string, string> UserNames = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly SemaphoreSlim NameLoadGate = new SemaphoreSlim(1, 1);
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>Разбирает авторов и заполняет кэш в фоне, один запрос на учетную запись за сеанс.</summary>
        public static async Task LoadNamesForCommentsAsync(IEnumerable<string> encodedComments)
        {
            string[] values = encodedComments.Where(value => !string.IsNullOrEmpty(value)).ToArray();
            if (values.Length == 0)
                return;

            await NameLoadGate.WaitAsync().ConfigureAwait(false);
            try
            {
                await Task.Run(() =>
                {
                    var users = values.SelectMany(value => value.Split(new[] { ClashesMainCollection.StringSeparatorItem }, StringSplitOptions.RemoveEmptyEntries))
                        .Select(value => value.Split(new[] { ClashesMainCollection.StringSeparatorSubItem }, StringSplitOptions.None))
                        .Where(parts => parts.Length == 3 && !string.IsNullOrWhiteSpace(parts[0]))
                        .Select(parts => parts[0]).Distinct(StringComparer.OrdinalIgnoreCase);
                    foreach (string user in users)
                    {
                        if (UserNames.ContainsKey(user))
                            continue;

                        string displayName = user;
                        try
                        {
                            DBUser dbUser = SQLiteMainService.SQLiteUserServiceInst.GetDBUser_ByUserName(user);
                            if (dbUser != null)
                            {
                                string fullName = (dbUser.Surname + " " + dbUser.Name).Trim();
                                if (!string.IsNullOrEmpty(fullName))
                                    displayName = fullName;
                            }
                        }
                        catch (Exception)
                        {
                            // Недоступный справочник не мешает показать сохраненный комментарий.
                        }
                        UserNames.TryAdd(user, displayName);
                    }
                }).ConfigureAwait(false);
            }
            finally
            {
                NameLoadGate.Release();
            }
        }

        public void RefreshUserName() => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(UserFullName)));

        public ReportItemComment(string message)
        {
            UserSystemName = SQLiteMainService.CurrentDBUser.SystemName;
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
                return !string.IsNullOrEmpty(UserSystemName) && UserNames.TryGetValue(UserSystemName, out string name)
                    ? name
                    : UserSystemName ?? string.Empty;
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
