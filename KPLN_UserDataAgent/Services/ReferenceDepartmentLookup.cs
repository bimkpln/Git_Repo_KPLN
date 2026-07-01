using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace KPLN_UserDataAgent.Services
{
    internal sealed class ReferenceDepartmentLookup
    {
        private const int BusyTimeoutMs = 30000;

        private readonly Dictionary<string, string> _departmentByUser =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public int UserCount
        {
            get { return _departmentByUser.Count; }
        }

        public static ReferenceDepartmentLookup Empty()
        {
            return new ReferenceDepartmentLookup();
        }

        public static ReferenceDepartmentLookup Load(string referenceDatabasePath)
        {
            ReferenceDepartmentLookup lookup = new ReferenceDepartmentLookup();
            if (string.IsNullOrWhiteSpace(referenceDatabasePath) || !File.Exists(referenceDatabasePath))
                return lookup;

            try
            {
                using (SQLiteConnection connection = OpenReadOnlyConnection(referenceDatabasePath))
                {
                    string tableName = ResolveTableName(connection, "Users", "User");
                    if (tableName == null)
                        return lookup;

                    HashSet<string> columns = GetTableColumns(connection, tableName);
                    string systemNameColumn = ResolveColumn(columns, "SystemName", "WindowsUser", "Login");
                    if (systemNameColumn == null)
                        return lookup;

                    string departmentIdColumn = ResolveColumn(
                        columns,
                        "SubDepartmentId",
                        "SubDepartmentID",
                        "SubDepartamentId",
                        "SubDepartamentID",
                        "DepartmentId",
                        "DepartmentID");
                    string departmentNameColumn = ResolveColumn(
                        columns,
                        "SubDepartament",
                        "SubDepartment",
                        "SubDepartmentName",
                        "Department",
                        "DepartmentName");

                    using (SQLiteCommand command = connection.CreateCommand())
                    {
                        command.CommandText =
                            "SELECT " +
                            QuoteIdentifier(systemNameColumn) + " AS SystemName, " +
                            SelectOptionalColumn(departmentIdColumn) + " AS DepartmentId, " +
                            SelectOptionalColumn(departmentNameColumn) + " AS DepartmentName " +
                            "FROM " + QuoteIdentifier(tableName) + ";";

                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string systemName = ReadString(reader, "SystemName");
                                if (string.IsNullOrWhiteSpace(systemName))
                                    continue;

                                lookup._departmentByUser[systemName] = ResolveDepartmentKey(
                                    ReadString(reader, "DepartmentId"),
                                    ReadString(reader, "DepartmentName"));
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            return lookup;
        }

        public static ReferenceDepartmentLookup LoadCache(string cachePath)
        {
            ReferenceDepartmentLookup lookup = new ReferenceDepartmentLookup();
            if (string.IsNullOrWhiteSpace(cachePath) || !File.Exists(cachePath))
                return lookup;

            try
            {
                foreach (string line in File.ReadAllLines(cachePath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    string[] parts = line.Split('\t');
                    if (parts.Length != 2)
                        continue;

                    string systemName = DecodeCacheValue(parts[0]);
                    string departmentKey = DecodeCacheValue(parts[1]);
                    if (!string.IsNullOrWhiteSpace(systemName))
                        lookup._departmentByUser[systemName] = departmentKey;
                }
            }
            catch
            {
            }

            return lookup;
        }

        public void SaveCache(string cachePath)
        {
            if (UserCount == 0 || string.IsNullOrWhiteSpace(cachePath))
                return;

            try
            {
                string directoryPath = Path.GetDirectoryName(cachePath);
                if (!string.IsNullOrWhiteSpace(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                string tempPath = cachePath + ".tmp";
                List<string> lines = _departmentByUser
                    .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(item => EncodeCacheValue(item.Key) + "\t" + EncodeCacheValue(item.Value))
                    .ToList();

                File.WriteAllLines(tempPath, lines.ToArray(), Encoding.UTF8);
                if (File.Exists(cachePath))
                    File.Delete(cachePath);

                File.Move(tempPath, cachePath);
            }
            catch
            {
            }
        }

        public string ResolveDepartmentKey(string windowsUser)
        {
            string departmentKey;
            if (!string.IsNullOrWhiteSpace(windowsUser)
                && _departmentByUser.TryGetValue(windowsUser, out departmentKey))
            {
                return CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentKey);
            }

            return CentralDatabasePathBuilder.UnknownDepartmentKey;
        }

        private static string ResolveDepartmentKey(string departmentId, string departmentName)
        {
            long id;
            if (long.TryParse(departmentId, NumberStyles.Integer, CultureInfo.InvariantCulture, out id) && id > 0)
                return "Department_" + id.ToString(CultureInfo.InvariantCulture);

            return CentralDatabasePathBuilder.NormalizeDepartmentKey(departmentName);
        }

        private static SQLiteConnection OpenReadOnlyConnection(string databasePath)
        {
            SQLiteConnection connection = new SQLiteConnection(
                "Data Source=" + databasePath + ";Version=3;Read Only=True;BusyTimeout=" + BusyTimeoutMs + ";Pooling=False;");
            connection.Open();
            ExecuteNonQuery(connection, "PRAGMA busy_timeout=" + BusyTimeoutMs + ";");
            return connection;
        }

        private static string ResolveTableName(SQLiteConnection connection, params string[] candidates)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = "SELECT name FROM sqlite_master WHERE type='table';";
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    List<string> tableNames = new List<string>();
                    while (reader.Read())
                    {
                        tableNames.Add(Convert.ToString(reader[0]));
                    }

                    return candidates.FirstOrDefault(
                        candidate => tableNames.Any(
                            tableName => string.Equals(tableName, candidate, StringComparison.OrdinalIgnoreCase)));
                }
            }
        }

        private static HashSet<string> GetTableColumns(SQLiteConnection connection, string tableName)
        {
            HashSet<string> columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = "PRAGMA table_info(" + QuoteIdentifier(tableName) + ");";
                using (SQLiteDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(Convert.ToString(reader["name"]));
                    }
                }
            }

            return columns;
        }

        private static string ResolveColumn(HashSet<string> columns, params string[] candidates)
        {
            return candidates.FirstOrDefault(columns.Contains);
        }

        private static string SelectOptionalColumn(string columnName)
        {
            return columnName == null ? "''" : QuoteIdentifier(columnName);
        }

        private static string QuoteIdentifier(string identifier)
        {
            return "\"" + (identifier ?? string.Empty).Replace("\"", "\"\"") + "\"";
        }

        private static string ReadString(SQLiteDataReader reader, string columnName)
        {
            object value = reader[columnName];
            return value == null || value == DBNull.Value ? string.Empty : Convert.ToString(value).Trim();
        }

        private static void ExecuteNonQuery(SQLiteConnection connection, string commandText)
        {
            using (SQLiteCommand command = connection.CreateCommand())
            {
                command.CommandText = commandText;
                command.ExecuteNonQuery();
            }
        }

        private static string EncodeCacheValue(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? string.Empty));
        }

        private static string DecodeCacheValue(string value)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value ?? string.Empty));
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    internal static class ReferenceDepartmentLookupService
    {
        private const int SuccessfulRefreshSeconds = 600;
        private const int FailedRetrySeconds = 60;

        private static readonly object Lock = new object();
        private static ReferenceDepartmentLookup _lookup = ReferenceDepartmentLookup.Empty();
        private static string _loadedPath = string.Empty;
        private static DateTime _lastLoadAttemptUtc = DateTime.MinValue;

        public static string ResolveDepartmentKey(string windowsUser, string referenceDatabasePath)
        {
            return ResolveDepartmentKey(windowsUser, referenceDatabasePath, false);
        }

        public static string ResolveDepartmentKey(string windowsUser, string referenceDatabasePath, bool forceRefresh)
        {
            EnsureLoaded(referenceDatabasePath, forceRefresh);
            return _lookup.ResolveDepartmentKey(windowsUser);
        }

        private static void EnsureLoaded(string referenceDatabasePath)
        {
            EnsureLoaded(referenceDatabasePath, false);
        }

        private static void EnsureLoaded(string referenceDatabasePath, bool forceRefresh)
        {
            string path = referenceDatabasePath ?? string.Empty;
            DateTime now = DateTime.UtcNow;

            lock (Lock)
            {
                int refreshSeconds = _lookup.UserCount > 0 ? SuccessfulRefreshSeconds : FailedRetrySeconds;
                bool samePath = string.Equals(_loadedPath, path, StringComparison.OrdinalIgnoreCase);
                if (!forceRefresh && samePath && (now - _lastLoadAttemptUtc).TotalSeconds < refreshSeconds)
                    return;

                ReferenceDepartmentLookup loadedLookup = ReferenceDepartmentLookup.Load(path);
                if (loadedLookup.UserCount > 0)
                {
                    _lookup = loadedLookup;
                    _lookup.SaveCache(ModuleData.LocalDepartmentLookupCachePath);
                }
                else if (_lookup.UserCount == 0)
                {
                    ReferenceDepartmentLookup cachedLookup =
                        ReferenceDepartmentLookup.LoadCache(ModuleData.LocalDepartmentLookupCachePath);
                    if (cachedLookup.UserCount > 0)
                        _lookup = cachedLookup;
                }

                _loadedPath = path;
                _lastLoadAttemptUtc = now;
            }
        }
    }
}