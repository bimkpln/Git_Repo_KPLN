using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using KPLN_OpeningHoleManager.Core;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    internal static class OpeningAuthorResolver
    {
        private static DBUser[] _users;
        private static DateTime _loadedAt;

        internal static string Resolve(string login)
        {
            if (string.IsNullOrWhiteSpace(login)) return OpeningHistoryRecord.UnknownAuthor;
            try
            {
                if (_users == null || DateTime.UtcNow - _loadedAt > TimeSpan.FromMinutes(5))
                {
                    _users = SQLiteMainService.SQLiteUserServiceInst.GetDBUsers().ToArray();
                    _loadedAt = DateTime.UtcNow;
                }
                return Resolve(login, _users);
            }
            catch
            {
                // Не повторять сетевой запрос для каждого отверстия при недоступной БД.
                _users = new DBUser[0];
                _loadedAt = DateTime.UtcNow;
                return login;
            }
        }

        internal static string Resolve(string login, IEnumerable<DBUser> users)
        {
            if (string.IsNullOrWhiteSpace(login)) return OpeningHistoryRecord.UnknownAuthor;
            DBUser[] matches = users.Where(u => Equal(u.RevitUserName, login)).ToArray();
            if (matches.Length == 0)
                matches = users.Where(u => Equal(u.SystemName, login)).ToArray();
            if (matches.Length == 0 && login.Contains("\\"))
                matches = users.Where(u => Equal(u.SystemName, login.Substring(login.LastIndexOf('\\') + 1))).ToArray();
            if (matches.Length != 1) return login;
            string name = (matches[0].Surname + " " + matches[0].Name).Trim();
            return name.Length == 0 ? login : name;
        }

        private static bool Equal(string left, string right) =>
            string.Equals(left?.Trim(), right?.Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
