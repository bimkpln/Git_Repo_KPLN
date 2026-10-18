using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;

namespace KPLN_Tools_OVVK.Common.OVVK_System
{
    /// <summary>Фактическое использование систем: от элементов модели к их системам.</summary>
    internal static class SystemUsageCollector
    {
        internal sealed class Usage
        {
            internal MEPSystem System { get; }
            internal HashSet<ElementId> ElementIds { get; } = new HashSet<ElementId>();
            internal ElementId RefreshCurveId { get; set; }

            internal Usage(MEPSystem system) { System = system; }
        }

        internal static Dictionary<ElementId, Usage> Collect(Document doc, Action checkCancellation = null)
        {
            var result = new Dictionary<ElementId, Usage>();
            var classes = new[] { typeof(Pipe), typeof(FlexPipe), typeof(Duct), typeof(FlexDuct),
                typeof(FamilyInstance), typeof(FabricationPart) };
            using (var filter = new ElementMulticlassFilter(classes))
            using (var collector = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType())
            {
                foreach (Element element in collector)
                {
                    checkCancellation?.Invoke();
                    if (!element.IsValidObject || element.Category == null)
                        continue;

                    ConnectorManager manager;
                    if (element is MEPCurve curve)
                    {
                        // Прямая ссылка участка подтверждает использование, даже при свободных концах.
                        Add(result, curve.MEPSystem, element, curve is Pipe || curve is Duct);
                        manager = curve.ConnectorManager;
                    }
                    else if (element is FamilyInstance family)
                        manager = family.MEPModel?.ConnectorManager;
                    else
                        manager = ((FabricationPart)element).ConnectorManager;

                    if (manager == null)
                        continue;
                    foreach (Connector connector in manager.Connectors)
                    {
                        // Оборудование может принадлежать нескольким системам через разные коннекторы.
                        if (!connector.IsValidObject || connector.ConnectorType == ConnectorType.Logical
                            || (connector.Domain != Domain.DomainHvac && connector.Domain != Domain.DomainPiping))
                            continue;
                        Add(result, connector.MEPSystem, element, false);
                    }
                }
            }
            return result;
        }

        private static void Add(Dictionary<ElementId, Usage> result, MEPSystem system, Element element, bool canRefresh)
        {
            if (system == null || !system.IsValidObject || (!(system is PipingSystem) && !(system is MechanicalSystem)))
                return;
            if (!result.TryGetValue(system.Id, out Usage usage))
            {
                usage = new Usage(system);
                result.Add(system.Id, usage);
            }
            usage.ElementIds.Add(element.Id);
            if (canRefresh && usage.RefreshCurveId == null)
                usage.RefreshCurveId = element.Id;
        }
    }
}
