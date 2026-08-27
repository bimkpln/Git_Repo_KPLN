using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Clashes
{
    public sealed class IosClasherService
    {
        public void Analyze(ClashAnalyzeRequest request)
        {
            if (request == null || !request.IsEnabled || request.CreatedElementIds.Count == 0)
                return;

            Outline analyzeOutline = BuildOutline(request.Doc, request.CreatedElementIds);

            // Корень интеграции с KPLN_IOSClasher:
            // сюда будет передаваться область analyzeOutline и набор новых элементов,
            // когда появится публичная точка входа или agreed reflection-contract чужого плагина.
        }

        private Outline BuildOutline(Document doc, IEnumerable<ElementId> elementIds)
        {
            XYZ min = null;
            XYZ max = null;

            foreach (ElementId id in elementIds)
            {
                Element element = doc.GetElement(id);
                BoundingBoxXYZ box = element?.get_BoundingBox(null);
                if (box == null)
                    continue;

                min = min == null
                    ? box.Min
                    : new XYZ(
                        System.Math.Min(min.X, box.Min.X),
                        System.Math.Min(min.Y, box.Min.Y),
                        System.Math.Min(min.Z, box.Min.Z));

                max = max == null
                    ? box.Max
                    : new XYZ(
                        System.Math.Max(max.X, box.Max.X),
                        System.Math.Max(max.Y, box.Max.Y),
                        System.Math.Max(max.Z, box.Max.Z));
            }

            if (min == null || max == null)
                return null;

            return new Outline(min, max);
        }
    }
}
