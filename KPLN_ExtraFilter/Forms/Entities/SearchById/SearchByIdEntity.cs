using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace KPLN_ExtraFilter.Forms.Entities.SearchById
{
    public sealed class SearchByIdEntity
    {
        public SearchByIdEntity(SearchByIdDocEntity sde, Element elem)
        {
            ElemDocEntity = sde;
            Elem = elem;
            
            Id = Elem.Id.ToString();
            CatName = Elem.Category != null ? Elem.Category.Name : "<Без категории>";

            // Определяю имя семейства
            if (Elem is Family family)
                ElemName = family.Name;
            else if (Elem is ElementType elType)
                ElemName = elType.FamilyName;
            else if (Elem is FamilyInstance fi)
                ElemName = fi.Symbol?.FamilyName;
            else
            {
                // "Семейство : Тип"
                string famAndType = Elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)?.AsValueString();
                if (!string.IsNullOrEmpty(famAndType))
                {
                    int idx = famAndType.IndexOf(':');
                    ElemName = idx > 0 ? famAndType.Substring(0, idx).Trim() : famAndType;
                }
                else
                    ElemName = "<Без имени>";
            }
        }

        public SearchByIdDocEntity ElemDocEntity { get; set; }

        public Element Elem { get; set; }

        public string Id { get; set; }

        public string CatName { get; set; }

        public string ElemName { get; set; }

        /// <summary>
        /// Получить BoundingBoxXYZ по элементу
        /// </summary>
        public BoundingBoxXYZ GetElemBBox()
        {
            BoundingBoxXYZ linkElemBBox = Elem.get_BoundingBox(null);
            if (linkElemBBox == null)
            {
                MessageBox.Show(
                    $"Выбранный элемент НЕ содержит 3D-геометрии, подобрать вид - невозможно.",
                    "Внимание",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return null;
            }
            
            Transform linkTrans = ElemDocEntity.SDE_RLI?.GetTotalTransform() ?? Transform.Identity;
            Transform boxTrans = linkElemBBox.Transform ?? Transform.Identity;
            IEnumerable<XYZ> hostCorners = GetBBoxCorners(linkElemBBox)
                .Select(p => linkTrans.OfPoint(boxTrans.OfPoint(p)));

            XYZ minPoint = null;
            XYZ maxPoint = null;
            foreach (XYZ point in hostCorners)
            {
                if (minPoint == null)
                {
                    minPoint = point;
                    maxPoint = point;
                    continue;
                }

                minPoint = new XYZ(
                    Math.Min(minPoint.X, point.X),
                    Math.Min(minPoint.Y, point.Y),
                    Math.Min(minPoint.Z, point.Z));

                maxPoint = new XYZ(
                    Math.Max(maxPoint.X, point.X),
                    Math.Max(maxPoint.Y, point.Y),
                    Math.Max(maxPoint.Z, point.Z));
            }

            return new BoundingBoxXYZ()
            {
                Min = minPoint,
                Max = maxPoint
            };
        }

        private static IEnumerable<XYZ> GetBBoxCorners(BoundingBoxXYZ bbox)
        {
            yield return new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z);
            yield return new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z);
            yield return new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z);
        }
    }
}
