using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ViewsAndLists_CopyByLevels.Services
{
    internal static class RevitHelpers
    {
        internal static bool TemplateControls(Document doc, View view, BuiltInParameter parameter)
        {
            View template = doc.GetElement(view.ViewTemplateId) as View;
            var id = new ElementId(parameter);
            return template != null && template.GetTemplateParameterIds().Contains(id)
                && !template.GetNonControlledTemplateParameterIds().Contains(id);
        }

        internal static string UniqueViewName(Document doc, string prefix)
        {
            var names = new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(View)).Cast<View>()
                .Select(v => v.Name), StringComparer.OrdinalIgnoreCase);
            string result = prefix;
            int suffix = 2;
            while (names.Contains(result)) result = prefix + " (" + suffix++ + ")";
            return result;
        }

        internal static void CopyParameters(Element source, Element target, ISet<ElementId> excluded, List<string> warnings)
        {
            foreach (Parameter parameter in source.Parameters)
            {
                if (excluded.Contains(parameter.Id) || !parameter.HasValue) continue;
                Parameter destination = parameter.IsShared ? target.get_Parameter(parameter.GUID) : target.get_Parameter(parameter.Definition);
                if (destination == null || destination.IsReadOnly || destination.StorageType != parameter.StorageType) continue;
                try
                {
                    switch (parameter.StorageType)
                    {
                        case StorageType.String: destination.Set(parameter.AsString() ?? ""); break;
                        case StorageType.Integer: destination.Set(parameter.AsInteger()); break;
                        case StorageType.Double: destination.Set(parameter.AsDouble()); break;
                        case StorageType.ElementId: destination.Set(parameter.AsElementId()); break;
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add("Параметр «" + parameter.Definition.Name + "» не перенесён: " + ex.Message);
                }
            }
        }

        internal static string InvalidName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "Пустое название";
            return name.IndexOfAny(new[] { '\\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '`', '~' }) >= 0
                ? "В названии есть запрещённые символы: " + name : null;
        }
    }
}
