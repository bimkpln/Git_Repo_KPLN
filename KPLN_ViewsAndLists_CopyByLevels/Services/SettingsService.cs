using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using KPLN_ViewsAndLists_Ribbon.Services.Schedules;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using KPLN_ViewsAndLists_CopyByLevels.Forms.ViewModels;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace KPLN_ViewsAndLists_CopyByLevels.Services
{
    internal static class SettingsService
    {
        private static string FileName(CopyByLevelsVM vm)
        {
            var model = vm.Model;
            string key = model.Document.PathName + "|" + model.Document.Title + "|" + model.SourceSheet.UniqueId;
            string hash;
            using (SHA256 sha = SHA256.Create()) hash = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(key))).Replace("-", "");
            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KPLN", ModuleData.ModuleName, "Settings");
            return Path.Combine(folder, hash + ".xml");
        }

        internal static void Save(CopyByLevelsVM vm)
        {
            var root = new XElement("Settings", new XAttribute("Version", 1),
                new XElement("Gap", vm.Gap), new XElement("NumberPrefix", vm.NumberPrefix), new XElement("StartNumber", vm.StartNumber),
                new XElement("SheetPattern", vm.SheetPattern), new XElement("ViewPattern", vm.ViewPattern),
                new XElement("MainPlan", vm.Model.MainPlan?.UniqueId ?? ""),
                new XElement("CopyPlanDetails", vm.Model.CopyPlanDetails),
                new XElement("CopySheetAnnotations", vm.Model.CopySheetAnnotations),
                new XElement("SkipExisting", vm.Model.SkipExisting));
            root.Add(new XElement("Contents", vm.Model.Contents.Select(c => new XElement("View", new XAttribute("Id", c.View.UniqueId),
                new XAttribute("Action", c.Action), new XAttribute("Replacement", c.Replacement == null ? "" : vm.Model.Document.GetElement(c.Replacement.Id).UniqueId)))));
            root.Add(new XElement("Levels", vm.Model.Levels.Select(l => new XElement("Level", new XAttribute("Id", l.Level.UniqueId),
                new XAttribute("Selected", l.Selected), new XAttribute("Floor", l.Floor ?? ""), new XAttribute("Group", l.Group ?? "")))));
            var schedules = new XElement("Schedules");
            foreach (ScheduleRow schedule in vm.Model.Schedules)
            {
                var saved = new XElement("Schedule", new XAttribute("Id", schedule.Schedule.UniqueId),
                    new XAttribute("Copy", schedule.Copy), new XAttribute("Include", schedule.Include));
                foreach (FilterRule filter in schedule.Filters)
                {
                    saved.Add(new XElement("Filter", new XAttribute("Index", filter.Index), new XAttribute("Identity", filter.FieldIdentity),
                        new XAttribute("Mode", filter.Mode), new XAttribute("Pattern", filter.Pattern ?? ""),
                        filter.Overrides.Select(o => new XElement("Value", new XAttribute("Level", o.Key), o.Value ?? ""))));
                }
                saved.Add(ScheduleLevelFilterService.Save(schedule));
                schedules.Add(saved);
            }
            root.Add(schedules);
            string path = FileName(vm);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            new XDocument(root).Save(path);
        }

        internal static void Load(CopyByLevelsVM vm)
        {
            string path = FileName(vm);
            if (!File.Exists(path)) return;
            try
            {
                XElement root = XDocument.Load(path).Root;
                if ((int?)root.Attribute("Version") != 1) return;
                vm.Gap = (string)root.Element("Gap") ?? vm.Gap;
                vm.NumberPrefix = (string)root.Element("NumberPrefix") ?? vm.NumberPrefix;
                vm.StartNumber = (string)root.Element("StartNumber") ?? vm.StartNumber;
                vm.SheetPattern = (string)root.Element("SheetPattern") ?? vm.SheetPattern;
                vm.ViewPattern = (string)root.Element("ViewPattern") ?? vm.ViewPattern;
                var mainPlan = vm.Plans.FirstOrDefault(p => vm.Model.Document.GetElement(p.Id).UniqueId == (string)root.Element("MainPlan"));
                if (mainPlan != null) vm.SelectedPlan = mainPlan;
                vm.Model.CopyPlanDetails = (bool?)root.Element("CopyPlanDetails") ?? true;
                vm.Model.CopySheetAnnotations = (bool?)root.Element("CopySheetAnnotations") ?? true;
                vm.Model.SkipExisting = (bool?)root.Element("SkipExisting") ?? true;
                foreach (XElement saved in root.Element("Contents")?.Elements("View") ?? Enumerable.Empty<XElement>())
                {
                    ContentRow content = vm.Model.Contents.FirstOrDefault(c => c.View.UniqueId == (string)saved.Attribute("Id"));
                    if (content == null) continue;
                    string action = (string)saved.Attribute("Action");
                    if (content.Actions.Contains(action)) content.Action = action;
                    content.Replacement = content.Replacements.FirstOrDefault(c => vm.Model.Document.GetElement(c.Id).UniqueId == (string)saved.Attribute("Replacement"));
                }
                foreach (XElement saved in root.Element("Levels")?.Elements("Level") ?? Enumerable.Empty<XElement>())
                {
                    LevelRow level = vm.Model.Levels.FirstOrDefault(l => l.Level.UniqueId == (string)saved.Attribute("Id"));
                    if (level == null) continue;
                    level.Selected = (bool?)saved.Attribute("Selected") ?? false;
                    level.Floor = (string)saved.Attribute("Floor") ?? level.Floor;
                    level.Group = (string)saved.Attribute("Group") ?? "";
                }
                foreach (XElement saved in root.Element("Schedules")?.Elements("Schedule") ?? Enumerable.Empty<XElement>())
                {
                    ScheduleRow schedule = vm.Model.Schedules.FirstOrDefault(s => s.Schedule.UniqueId == (string)saved.Attribute("Id"));
                    if (schedule == null) continue;
                    schedule.Copy = (bool?)saved.Attribute("Copy") ?? true;
                    schedule.Include = (bool?)saved.Attribute("Include") ?? true;
                    foreach (XElement setting in saved.Elements("Filter"))
                    {
                        FilterRule filter = schedule.Filters.FirstOrDefault(f => f.Index == (int?)setting.Attribute("Index") && f.FieldIdentity == (string)setting.Attribute("Identity"));
                        if (filter == null) continue;
                        string mode = (string)setting.Attribute("Mode");
                        if (!FilterModes.All.Contains(mode)) continue;
                        filter.Mode = mode;
                        filter.Pattern = (string)setting.Attribute("Pattern") ?? "{Этаж}";
                        foreach (XElement value in setting.Elements("Value")) filter.Overrides[(string)value.Attribute("Level")] = value.Value;
                    }
                    ScheduleLevelFilterService.Load(vm.Model.Document, schedule, saved.Element("LevelFilters"));
                }
            }
            catch (Exception)
            {
                // Повреждённая или устаревшая настройка не должна мешать открыть инструмент.
            }
        }
    }
}
