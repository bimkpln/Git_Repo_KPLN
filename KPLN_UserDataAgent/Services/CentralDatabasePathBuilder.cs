using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace KPLN_UserDataAgent.Services
{
    internal static class CentralDatabasePathBuilder
    {
        public const string UnknownDepartmentKey = "UnknownDepartment";
        private const int MaxDepartmentKeyLength = 80;

        public static string GetDatabasePath(string centralRootPath, string departmentKey, string recordTime)
        {
            return GetDatabasePath(centralRootPath, departmentKey, recordTime, "KPLN_UserDataAgent");
        }

        public static string GetDatabasePath(
            string centralRootPath,
            string departmentKey,
            string recordTime,
            string databaseNamePrefix)
        {
            string safeDepartmentKey = NormalizeDepartmentKey(departmentKey);
            string month = ParseRecordTime(recordTime).ToString("yyyy-MM", CultureInfo.InvariantCulture);
            string safeDatabaseNamePrefix = NormalizeDatabaseNamePrefix(databaseNamePrefix);
            string fileName = string.Format(
                CultureInfo.InvariantCulture,
                "{0}_{1}_{2}.db",
                safeDatabaseNamePrefix,
                safeDepartmentKey,
                month);

            return Path.Combine(centralRootPath, safeDepartmentKey, month, fileName);
        }

        public static string NormalizeDepartmentKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return UnknownDepartmentKey;

            char[] invalidChars = Path.GetInvalidFileNameChars();
            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char character in value.Trim())
            {
                if (Array.IndexOf(invalidChars, character) >= 0 || char.IsControl(character))
                {
                    builder.Append('_');
                    continue;
                }

                builder.Append(char.IsWhiteSpace(character) ? '_' : character);
            }

            string result = builder.ToString().Trim('_', '.', ' ');
            if (string.IsNullOrWhiteSpace(result))
                return UnknownDepartmentKey;

            return result.Length <= MaxDepartmentKeyLength
                ? result
                : result.Substring(0, MaxDepartmentKeyLength).Trim('_', '.', ' ');
        }

        public static string NormalizeDatabaseNamePrefix(string value)
        {
            string result = NormalizeDepartmentKey(value);
            return string.Equals(result, UnknownDepartmentKey, StringComparison.OrdinalIgnoreCase)
                ? "KPLN_UserDataAgent"
                : result;
        }

        public static bool IsDatabaseFilePath(string path)
        {
            string extension = Path.GetExtension(path ?? string.Empty);
            return string.Equals(extension, ".db", StringComparison.OrdinalIgnoreCase)
                || string.Equals(extension, ".sqlite", StringComparison.OrdinalIgnoreCase);
        }

        public static bool TryParseMonthDirectoryName(string value, out DateTime month)
        {
            return DateTime.TryParseExact(
                value ?? string.Empty,
                "yyyy-MM",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out month);
        }

        public static DateTime ParseRecordTime(string value)
        {
            DateTime parsed;
            return DateTime.TryParseExact(
                value ?? string.Empty,
                "yyyy.MM.dd. HH:mm:ss",
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeLocal,
                out parsed)
                ? parsed
                : DateTime.Now;
        }
    }
}