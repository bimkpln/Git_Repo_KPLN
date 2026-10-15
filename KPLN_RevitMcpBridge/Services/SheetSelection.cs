using Autodesk.Revit.DB;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;

namespace KPLN_RevitMcpBridge.Services
{
    internal static class SheetSelection
    {
        internal static ViewSheet[] Resolve(Document doc, IDictionary<string, object> input)
        {
            var result = new List<ViewSheet>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            // Проверяем весь заданный состав до экспорта. Нет неявного перехода на все листы.
            foreach (var value in input.List("sheet_unique_ids"))
            {
                var id = value as string;
                if (string.IsNullOrWhiteSpace(id) || !seen.Add(id))
                    throw new BridgeException("invalid_sheet_selection", "Ожидаются непустые уникальные UniqueId листов.");

                var sheet = doc.GetElement(id) as ViewSheet;
                if (sheet == null || sheet.IsPlaceholder || !sheet.CanBePrinted)
                    throw new BridgeException("invalid_sheet_selection", "Лист не найден в активном документе или недоступен для печати: " + id);

                result.Add(sheet);
            }

            return result.ToArray();
        }
    }
}
