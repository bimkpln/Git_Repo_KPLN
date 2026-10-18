using Autodesk.Revit.DB;
using KPLN_Library_ConfigWorker;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
using KPLN_Tools_OVVK.Forms.Entities;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    internal static class SystemProjectConfig
    {
        internal static string ModelName(Document doc)
        {
            string path = SQLiteDocService.GetFileFullName(doc);
            return string.IsNullOrEmpty(path) ? doc.Title : Path.GetFileNameWithoutExtension(path.Replace('/', '\\'));
        }

        private static string Folder(Document doc)
        {
            if (!doc.IsWorkshared || doc.IsDetached)
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KPLN_Temp", "KPLN_Config");

            string path = SQLiteDocService.GetFileFullName(doc);
            if (!path.StartsWith("RSN:", StringComparison.OrdinalIgnoreCase))
                return Path.Combine(Path.GetDirectoryName(path), "KPLN_Config");

            var project = SQLiteMainService.SQLitePrjServiceInst.GetDBProject_ByRevitDocFileNameANDRVersion(
                path, int.Parse(doc.Application.VersionNumber));
            if (project == null || string.IsNullOrWhiteSpace(project.Code))
                throw new InvalidOperationException("Не найден проект Revit Server для общего конфига.");
            return Path.Combine(@"Z:\KPLN_Temp\KPLN_Config", project.Code);
        }

        internal static JObject Model(JObject root, string name)
        {
            var models = (JObject)root["Models"];
            var property = models.Properties().FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
                return (JObject)property.Value;
            var model = new JObject();
            models[name] = model;
            return model;
        }

        internal static void Access(Document doc, Func<JObject, bool> update)
        {
            string folder = Folder(doc);
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "OVVK_Systems.json");
            // Все разделы изменяются под одной блокировкой с перечитыванием файла.
            using (var guard = new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                string json = File.Exists(path) ? File.ReadAllText(path, Encoding.UTF8) : null;
                bool created = string.IsNullOrWhiteSpace(json);
                JObject root = created ? new JObject { ["Version"] = 1, ["Settings"] = new JObject(), ["Models"] = new JObject() } : JObject.Parse(json);
                if ((int?)root["Version"] != 1 || !(root["Settings"] is JObject) || !(root["Models"] is JObject))
                    throw new InvalidDataException("Повреждён общий конфиг систем.");

                if (created)
                    ImportSettings(doc, root);
                bool imported = ImportRefresh(doc, root, folder);
                if (update(root) || created || imported)
                    Write(path, root);
            }
        }

        private static void ImportSettings(Document doc, JObject root)
        {
            var type = doc.IsWorkshared && !doc.IsDetached ? ConfigType.Shared : ConfigType.Local;
            int version = int.Parse(doc.Application.VersionNumber);
            var settings = ConfigService.ReadConfigFile<OVVK_SystemManagerM>(version, doc, type, "OVVK_SystemManager") as OVVK_SystemManagerM;
            if (settings != null)
                root["Settings"] = JObject.FromObject(settings.ToJson());
            var merge = ConfigService.ReadConfigFile<OVVK_MergeSystemsConfigM>(version, doc, type, "OVVK_MergeSystems") as OVVK_MergeSystemsConfigM;
            foreach (var entry in merge?.Models ?? new System.Collections.Generic.List<OVVK_MergeSystemsConfigM.ModelSettings>())
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.ModelName))
                    continue;
                JObject model = Model(root, entry.ModelName);
                var groups = model["MergeSystems"] as JArray ?? new JArray();
                model["MergeSystems"] = groups;
                groups.Add(JObject.FromObject(entry));
            }
        }

        private static bool ImportRefresh(Document doc, JObject root, string folder)
        {
            JObject model = Model(root, ModelName(doc));
            if (model["Refresh"] != null)
                return false;

            ModelPath central = doc.IsWorkshared ? doc.GetWorksharingCentralModelPath() : null;
            string modelPath = (central == null ? doc.PathName : ModelPathUtils.ConvertModelPathToUserVisiblePath(central)) ?? string.Empty;
            string identity = doc.Application.VersionNumber + "|" + (string.IsNullOrEmpty(modelPath)
                ? doc.ProjectInformation.UniqueId : modelPath.ToUpperInvariant());
            string oldFolder = central == null && !string.IsNullOrEmpty(modelPath)
                ? Path.Combine(Path.GetDirectoryName(modelPath), "KPLN_Config") : folder;
            using (var hash = SHA256.Create())
            {
                string key = BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(identity))).Replace("-", "");
                string oldPath = Path.Combine(oldFolder, "OVVK_SystemNameRefresh_" + key + ".json");
                JObject state = File.Exists(oldPath) ? JObject.Parse(File.ReadAllText(oldPath, Encoding.UTF8))
                    : new JObject { ["Required"] = false };
                if ((bool?)state["Required"] == true && string.IsNullOrEmpty((string)state["Token"]))
                    throw new InvalidDataException("Повреждена старая метка обновления систем.");
                model["Refresh"] = state;
            }
            return true;
        }

        internal static OVVK_MergeSystemsConfigM ReadMerge(Document doc)
        {
            var result = new OVVK_MergeSystemsConfigM();
            Access(doc, root =>
            {
                foreach (var property in ((JObject)root["Models"]).Properties())
                    foreach (var item in (property.Value["MergeSystems"] as JArray) ?? new JArray())
                    {
                        var settings = item.ToObject<OVVK_MergeSystemsConfigM.ModelSettings>();
                        settings.ModelName = property.Name;
                        result.Models.Add(settings);
                    }
                return false;
            });
            return result;
        }

        internal static void SaveMerge(Document doc, OVVK_MergeSystemsConfigM.ModelSettings settings)
        {
            Access(doc, root =>
            {
                JObject model = Model(root, ModelName(doc));
                var groups = model["MergeSystems"] as JArray ?? new JArray();
                foreach (var entry in groups.Where(item => (string)item["ParameterName"] == settings.ParameterName).ToArray())
                    entry.Remove();
                var saved = JObject.FromObject(settings);
                saved.Remove("ModelName");
                groups.Add(saved);
                model["MergeSystems"] = groups;
                return true;
            });
        }

        private static void Write(string path, JObject root)
        {
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                File.WriteAllText(temporary, root.ToString(Formatting.Indented), Encoding.UTF8);
                if (File.Exists(path))
                    File.Replace(temporary, path, null);
                else
                    File.Move(temporary, path);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }
    }
}
