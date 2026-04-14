using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Clashes_Ribbon.Tools
{
    /// <summary>
    /// Выделение элементов в модели
    /// </summary>
    internal static class SelectInDocTools
    {
        internal static void SelectElemsInDoc(UIDocument uidoc, List<ElementId> elemsId)
        {
            Element[] elems = elemsId
                .Select(id => uidoc.Document.GetElement(id))
                .OfType<Element>()
                .ToArray();

            // Анализ на изоляцию ИОС. Если да - добавляю и их в коллекцию
            InsulationLiningBase[] mepInsBase = elems
                .Where(el => el is InsulationLiningBase)
                .OfType<InsulationLiningBase>()
                .ToArray();
            if (mepInsBase.Length > 0)
                elemsId.AddRange(mepInsBase.Select(mib => mib.HostElementId));

            uidoc.Selection.SetElementIds(elemsId);
        }
    }
}
