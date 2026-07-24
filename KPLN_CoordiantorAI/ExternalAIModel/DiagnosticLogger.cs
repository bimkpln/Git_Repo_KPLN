using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_CoordiantorAI.ExternalAIModel
{
    internal class DiagnosticLogger
    {
        private const int RetentionDays = 3;
        private readonly string _logFilePath;
        private readonly object _syncRoot = new object();

        public DiagnosticLogger(string logFolder)
        {
            string diagnosticsFolder = GetDiagnosticsFolder(logFolder);
            if (!Directory.Exists(diagnosticsFolder))
                Directory.CreateDirectory(diagnosticsFolder);

            string userName = SanitizeFileName(Environment.UserName);
            CleanupOldLogs(diagnosticsFolder, userName);

            string fileName = string.Format(
                "{0}_diagnostic_{1}.txt",
                userName,
                DateTime.Now.ToString("yyyy-MM-dd"));

            _logFilePath = Path.Combine(diagnosticsFolder, fileName);
        }

        public void LogEvent(string requestId, string eventName, IDictionary<string, object> details = null)
        {
            try
            {
                StringBuilder line = new StringBuilder();
                line.Append(FormatDate(DateTime.Now));
                line.Append(" | requestId=");
                line.Append(string.IsNullOrWhiteSpace(requestId) ? "-" : requestId);
                line.Append(" | ");
                line.Append(eventName ?? "UNKNOWN");

                if (details != null)
                {
                    foreach (KeyValuePair<string, object> item in details)
                    {
                        line.Append(" | ");
                        line.Append(item.Key);
                        line.Append("=");
                        line.Append(NormalizeValue(item.Value));
                    }
                }

                AppendLine(line.ToString());
            }
            catch
            {
            }
        }

        public void LogException(string requestId, string eventName, Exception exception, IDictionary<string, object> details = null)
        {
            Dictionary<string, object> mergedDetails = new Dictionary<string, object>();
            if (details != null)
            {
                foreach (KeyValuePair<string, object> item in details)
                    mergedDetails[item.Key] = item.Value;
            }

            if (exception != null)
            {
                mergedDetails["exceptionType"] = exception.GetType().FullName;
                mergedDetails["message"] = exception.Message;
                mergedDetails["stackTrace"] = exception.ToString();
            }

            LogEvent(requestId, eventName, mergedDetails);
        }

        private void AppendLine(string line)
        {
            lock (_syncRoot)
            {
                File.AppendAllText(_logFilePath, line + Environment.NewLine, Encoding.UTF8);
            }
        }

        private static string GetDiagnosticsFolder(string logFolder)
        {
            logFolder = (logFolder ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(logFolder))
                return Path.Combine(logFolder, "Diagnostics");

            //формируется путь по примеру: C:\Users\mtarchokov\AppData\Local\KPLN\CoordinatorAI\Diagnostics 
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "KPLN",
                "CoordinatorAI",
                "Diagnostics");
        }

        private static string SanitizeFileName(string value)
        {
            string name = string.IsNullOrWhiteSpace(value) ? "user" : value.Trim();
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
                name = name.Replace(invalidChar, '_');

            return name;
        }

        private static void CleanupOldLogs(string diagnosticsFolder, string userName)
        {
            try
            {
                DateTime minDateToKeep = DateTime.Today.AddDays(-(RetentionDays - 1));
                string searchPattern = string.Format("{0}_diagnostic_*.txt", userName);
                foreach (string filePath in Directory.GetFiles(diagnosticsFolder, searchPattern))
                {
                    string fileName = Path.GetFileNameWithoutExtension(filePath);
                    string dateText = fileName.Substring((userName + "_diagnostic_").Length);
                    DateTime logDate;
                    if (!DateTime.TryParseExact(
                        dateText,
                        "yyyy-MM-dd",
                        System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None,
                        out logDate))
                    {
                        continue;
                    }

                    if (logDate.Date < minDateToKeep)
                        File.Delete(filePath);
                }
            }
            catch
            {
            }
        }


        private static string NormalizeValue(object value)
        {
            if (value == null)
                return "";

            string text = Convert.ToString(value);
            if (string.IsNullOrEmpty(text))
                return "";

            return text
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("|", "/");
        }

        private static string FormatDate(DateTime value)
        {
            return value.ToString("yyyy-MM-dd HH:mm:ss.fff");
        }

    }
}
