using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace RevitMcpStdioProxy
{
    internal sealed class RevitInstanceDiscovery
    {
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer();

        public string InstancesDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KPLN",
                    "CoordinatorAI",
                    "MCP",
                    "instances");
            }
        }

        public List<RevitInstanceInfo> Discover()
        {
            List<RevitInstanceInfo> result = new List<RevitInstanceInfo>();
            if (!Directory.Exists(InstancesDirectory))
                return result;

            foreach (string path in Directory.GetFiles(InstancesDirectory, "*.json"))
            {
                RevitInstanceInfo instance;
                if (TryReadLiveInstance(path, out instance))
                    result.Add(instance);
            }

            return result
                .OrderBy(i => i.RevitVersion)
                .ThenBy(i => i.ProcessStartTimeUtc)
                .ThenBy(i => i.ProcessId)
                .ToList();
        }

        private static bool TryReadLiveInstance(string path, out RevitInstanceInfo instance)
        {
            instance = null;

            try
            {
                Dictionary<string, object> data =
                    Serializer.DeserializeObject(File.ReadAllText(path)) as Dictionary<string, object>;
                if (data == null)
                {
                    DeleteStaleRegistration(path);
                    return false;
                }

                RevitInstanceInfo candidate = new RevitInstanceInfo
                {
                    InstanceId = GetString(data, "instance_id"),
                    ProcessId = GetInt(data, "process_id"),
                    RevitVersion = GetInt(data, "revit_version"),
                    Endpoint = GetString(data, "endpoint"),
                    PluginVersion = GetString(data, "plugin_version"),
                    ProcessStartTimeUtc = GetDateTime(data, "process_start_time_utc")
                };

                if (string.IsNullOrWhiteSpace(candidate.InstanceId)
                    || candidate.ProcessId <= 0
                    || candidate.RevitVersion <= 0
                    || !IsValidLocalEndpoint(candidate.Endpoint))
                {
                    DeleteStaleRegistration(path);
                    return false;
                }

                using (Process process = Process.GetProcessById(candidate.ProcessId))
                {
                    if (process.HasExited
                        || !string.Equals(process.ProcessName, "Revit", StringComparison.OrdinalIgnoreCase))
                    {
                        DeleteStaleRegistration(path);
                        return false;
                    }

                    DateTime actualStartTimeUtc = process.StartTime.ToUniversalTime();
                    if (candidate.ProcessStartTimeUtc != DateTime.MinValue
                        && Math.Abs((actualStartTimeUtc - candidate.ProcessStartTimeUtc).TotalSeconds) > 5)
                    {
                        DeleteStaleRegistration(path);
                        return false;
                    }

                    candidate.ProcessStartTimeUtc = actualStartTimeUtc;
                    candidate.WindowTitle = process.MainWindowTitle ?? string.Empty;
                    candidate.ModelName = ExtractModelName(candidate.WindowTitle);
                }

                instance = candidate;
                return true;
            }
            catch (ArgumentException)
            {
                DeleteStaleRegistration(path);
                return false;
            }
            catch (InvalidOperationException)
            {
                DeleteStaleRegistration(path);
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsValidLocalEndpoint(string endpoint)
        {
            Uri uri;
            if (!Uri.TryCreate(endpoint, UriKind.Absolute, out uri))
                return false;

            return string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                && (string.Equals(uri.Host, "127.0.0.1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(uri.Host, "localhost", StringComparison.OrdinalIgnoreCase));
        }

        private static string ExtractModelName(string windowTitle)
        {
            if (string.IsNullOrWhiteSpace(windowTitle))
                return string.Empty;

            int openBracket = windowTitle.IndexOf('[');
            int closeBracket = windowTitle.LastIndexOf(']');
            if (openBracket < 0 || closeBracket <= openBracket)
                return string.Empty;

            string documentAndView = windowTitle.Substring(
                openBracket + 1,
                closeBracket - openBracket - 1);
            int separator = documentAndView.IndexOf(" - ", StringComparison.Ordinal);
            return separator < 0
                ? documentAndView.Trim()
                : documentAndView.Substring(0, separator).Trim();
        }

        private static string GetString(Dictionary<string, object> data, string name)
        {
            object value;
            return data.TryGetValue(name, out value) && value != null
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private static int GetInt(Dictionary<string, object> data, string name)
        {
            object value;
            if (!data.TryGetValue(name, out value) || value == null)
                return 0;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static DateTime GetDateTime(Dictionary<string, object> data, string name)
        {
            string value = GetString(data, name);
            DateTime result;
            return DateTime.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out result)
                ? result.ToUniversalTime()
                : DateTime.MinValue;
        }

        private static void DeleteStaleRegistration(string path)
        {
            try
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
            }
        }
    }

    internal sealed class RevitInstanceInfo
    {
        public string InstanceId { get; set; }
        public int ProcessId { get; set; }
        public int RevitVersion { get; set; }
        public string Endpoint { get; set; }
        public string PluginVersion { get; set; }
        public DateTime ProcessStartTimeUtc { get; set; }
        public string WindowTitle { get; set; }
        public string ModelName { get; set; }
    }
}
