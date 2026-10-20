using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using KPLN_ViewsAndLists_Ribbon.Services.Schedules;
using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_CopyByLevels.Common;
using KPLN_ViewsAndLists_CopyByLevels.Forms.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using FilterRule = KPLN_ViewsAndLists_Ribbon.Common.Schedules.FilterRule;

namespace KPLN_ViewsAndLists_CopyByLevels.Services
{
    internal static class SourceReader
    {
        internal static List<ElementChoice> Sheets(Document doc)
        {
            return new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).Cast<ViewSheet>()
                .Where(s => !s.IsPlaceholder).OrderBy(s => s.SheetNumber)
                .Select(s => new ElementChoice { Id = s.Id, Name = s.SheetNumber + " — " + s.Name }).ToList();
        }

        internal static CopyByLevelsM Read(Document doc, ViewSheet sheet)
        {
            var model = new CopyByLevelsM { Document = doc, SourceSheet = sheet };
            var views = new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                .Where(v => !v.IsTemplate && !(v is ViewSheet) && !(v is ViewSchedule)).OrderBy(v => v.Name).ToList();
            var replacements = views.Select(v => new ElementChoice { Id = v.Id, Name = v.Name }).ToList();

            foreach (Viewport port in sheet.GetAllViewports().Select(id => doc.GetElement(id)).OfType<Viewport>())
            {
                var view = (View)doc.GetElement(port.ViewId);
                var plan = view as ViewPlan;
                bool eligible = plan != null && plan.GenLevel != null && plan.ViewType != ViewType.AreaPlan;
                var actions = new List<string>();
                if (eligible) actions.Add(CopyActions.NewPlan);
                if (view.ViewType == ViewType.Legend) actions.Add(CopyActions.Shared);
                if (view.CanViewBeDuplicated(ViewDuplicateOption.WithDetailing)) actions.Add(CopyActions.Duplicate);
                actions.Add(CopyActions.Replace);
                actions.Add(CopyActions.Skip);

                model.Contents.Add(new ContentRow
                {
                    Viewport = port, View = view, Actions = actions, Replacements = replacements,
                    Action = eligible ? CopyActions.NewPlan : actions[0]
                });
            }

            // Наиболее крупное окно плана — предварительный выбор, пользователь может изменить его.
            model.MainPlan = model.Contents.Where(c => c.Action == CopyActions.NewPlan)
                .OrderByDescending(c =>
                {
                    Outline box = c.Viewport.GetBoxOutline();
                    return (box.MaximumPoint.X - box.MinimumPoint.X) * (box.MaximumPoint.Y - box.MinimumPoint.Y);
                }).Select(c => c.View as ViewPlan).FirstOrDefault();
            foreach (ContentRow content in model.Contents.Where(c => c.Action == CopyActions.NewPlan && c.View.Id != model.MainPlan?.Id))
                content.Action = content.Actions.Contains(CopyActions.Duplicate) ? CopyActions.Duplicate : CopyActions.Skip;

            var placedSchedules = new FilteredElementCollector(doc).OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>().Where(i => i.OwnerViewId == sheet.Id && !i.IsTitleblockRevisionSchedule);
            foreach (var group in placedSchedules.GroupBy(i => i.ScheduleId))
            {
                var schedule = doc.GetElement(group.Key) as ViewSchedule;
                if (schedule == null) continue;
                var row = new ScheduleRow { Schedule = schedule, Instances = group.ToList() };
                ScheduleService.ReadFilters(doc, row);
                model.Schedules.Add(row);
            }

            PopulateLevels(model);
            return model;
        }

        internal static void PopulateLevels(CopyByLevelsM model)
        {
            model.Levels.Clear();
            if (model.MainPlan?.GenLevel == null) return;
            var plans = new FilteredElementCollector(model.Document).OfClass(typeof(ViewPlan)).Cast<ViewPlan>()
                .Where(v => !v.IsTemplate && v.GenLevel != null).ToList();
            var levels = new FilteredElementCollector(model.Document).OfClass(typeof(Level)).Cast<Level>()
                .OrderBy(l => l.ProjectElevation).ThenBy(l => l.Name).ToList();
            int floor = 0;
            int baseFloor = int.Parse(NamingRules.GuessFloor(model.MainPlan.GenLevel.Name, 0));
            foreach (Level level in levels)
            {
                if (level.ProjectElevation <= model.MainPlan.GenLevel.ProjectElevation + 1e-6) continue;
                floor++;
                model.Levels.Add(new LevelRow
                {
                    Level = level, Floor = NamingRules.GuessFloor(level.Name, baseFloor + floor),
                    Existing = string.Join("; ", plans.Where(p => p.GenLevel.Id == level.Id).Select(p => p.Name)),
                    Hint = levels.Count(l => Math.Abs(l.ProjectElevation - level.ProjectElevation) < 1.0 / 304.8) > 1
                        ? "Несколько уровней на этой отметке — проверьте корпус и этаж"
                        : NamingRules.GuessFloor(level.Name, int.MinValue) == int.MinValue.ToString()
                            ? "Обозначение этажа предположено — проверьте" : ""
                });
            }
        }
    }
}
