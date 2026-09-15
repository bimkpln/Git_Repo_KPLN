using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace KPLN_CoordiantorAI.ExternalAIModel.Mcp
{
    internal static class RevitMcpInstanceRegistry
    {
        private const string InstancesFolderName = "instances";

        public static string InstancesDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "KPLN",
                    "CoordinatorAI",
                    "MCP",
                    InstancesFolderName);
            }
        }

        public static void Register(string instanceId, string endpoint, int revitVersion)
        {
            if (string.IsNullOrWhiteSpace(instanceId))
                throw new ArgumentException("Instance id is required.", "instanceId");
            if (string.IsNullOrWhiteSpace(endpoint))
                throw new ArgumentException("Endpoint is required.", "endpoint");

            Directory.CreateDirectory(InstancesDirectory);

            Process process = Process.GetCurrentProcess();
            JObject registration = new JObject
            {
                { "instance_id", instanceId },
                { "process_id", process.Id },
                { "process_start_time_utc", process.StartTime.ToUniversalTime().ToString("o") },
                { "revit_version", revitVersion },
                { "endpoint", endpoint },
                { "plugin_version", Assembly.GetExecutingAssembly().GetName().Version.ToString() },
                { "registered_at_utc", DateTime.UtcNow.ToString("o") }
            };

            string targetPath = GetRegistrationPath(instanceId);
            string temporaryPath = targetPath + "." + Guid.NewGuid().ToString("N") + ".tmp";

            try
            {
                File.WriteAllText(
                    temporaryPath,
                    JsonConvert.SerializeObject(registration, Formatting.Indented),
                    new UTF8Encoding(false));

                if (File.Exists(targetPath))
                    File.Delete(targetPath);

                File.Move(temporaryPath, targetPath);
            }
            finally
            {
                try
                {
                    if (File.Exists(temporaryPath))
                        File.Delete(temporaryPath);
                }
                catch
                {
                }
            }
        }

        public static void Unregister(string instanceId)
        {
            if (string.IsNullOrWhiteSpace(instanceId))
                return;

            string path = GetRegistrationPath(instanceId);
            if (File.Exists(path))
                File.Delete(path);
        }

        private static string GetRegistrationPath(string instanceId)
        {
            return Path.Combine(InstancesDirectory, instanceId + ".json");
        }
    }
}
