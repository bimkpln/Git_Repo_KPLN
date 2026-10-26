using Autodesk.Revit.DB;
using System;

namespace KPLN_RevitMcpBridge.Services
{
    /// <summary>Содержимое ручных аннотаций без изменения документа/семейств.</summary>
    internal static class AnnotationContentService
    {
        internal static bool Supports(Element element)
        {
            return element is TextNote || (element is FamilyInstance && element.Category != null
                && RevitService.IdValue(element.Category.Id) == (long)BuiltInCategory.OST_GenericAnnotation);
        }

        internal static object Read(Document doc, Element element)
        {
            if (!Supports(element)) return null;

            var note = element as TextNote;
            var type = doc.GetElement(element.GetTypeId());
            var owner = doc.GetElement(element.OwnerViewId) as View;
            return new
            {
                kind = note != null ? "text_note" : "generic_annotation",
                element = RevitService.ElementInfo(element),
                type = type == null ? null : RevitService.ElementInfo(type),
                owner_view_id = RevitService.IdValue(element.OwnerViewId).ToString(),
                text = note?.Text,
                text_status = note != null ? "available" : "not_resolved_from_family_labels",
                family_name = (type as FamilySymbol)?.FamilyName,
                parameters = ParameterService.Read(element),
                type_parameters = type == null ? null : ParameterService.Read(type),
                location = note != null ? new { kind = "point", point_mm = Point(note.Coord) } : ElementLocationService.Read(element),
                bounds = Bounds(element, owner),
                association_semantics = "standalone_annotation; no model tag target; parameter values do not prove rendered label content",
                visibility_semantics = "owner_view_elements; crop, visibility and occlusion not evaluated"
            };
        }

        private static double[] Point(XYZ p) => new[] { p.X * 304.8, p.Y * 304.8, p.Z * 304.8 };

        private static object Bounds(Element element, View view)
        {
            try
            {
                var box = element.get_BoundingBox(view);
                if (box == null) return new { status = "unavailable" };

                // Сохраняем базис: min/max заданы в локальной системе BoundingBoxXYZ.
                return new
                {
                    status = "available", units = "mm", coordinate_system = "box_local",
                    min_mm = Point(box.Min), max_mm = Point(box.Max), origin_mm = Point(box.Transform.Origin),
                    basis_x = new[] { box.Transform.BasisX.X, box.Transform.BasisX.Y, box.Transform.BasisX.Z },
                    basis_y = new[] { box.Transform.BasisY.X, box.Transform.BasisY.Y, box.Transform.BasisY.Z },
                    basis_z = new[] { box.Transform.BasisZ.X, box.Transform.BasisZ.Y, box.Transform.BasisZ.Z }
                };
            }
            catch (Exception ex)
            {
                return new { status = "unavailable", reason = ex.Message };
            }
        }
    }
}
