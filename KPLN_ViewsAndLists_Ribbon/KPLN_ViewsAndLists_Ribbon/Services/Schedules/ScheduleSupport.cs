using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ViewsAndLists_Ribbon.Services.Schedules
{
    public static class ScheduleSupport
    {
        public static bool TemplateControls(Document doc, View view, BuiltInParameter parameter)
        {
            View template = doc.GetElement(view.ViewTemplateId) as View;
            var id = new ElementId(parameter);
            return template != null && template.GetTemplateParameterIds().Contains(id)
                && !template.GetNonControlledTemplateParameterIds().Contains(id);
        }

        public static string UniqueViewName(Document doc, string prefix)
        {
            var names = new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                .Select(v => v.Name), StringComparer.OrdinalIgnoreCase);
            string result = prefix;
            int suffix = 2;
            while (names.Contains(result)) result = prefix + " (" + suffix++ + ")";
            return result;
        }

        public static string InvalidName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Пустое название";
            return name.IndexOfAny(new[] { '\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' }) >= 0
                ? "В названии есть запрещённые символы: " + name : null;
        }
    }
}
