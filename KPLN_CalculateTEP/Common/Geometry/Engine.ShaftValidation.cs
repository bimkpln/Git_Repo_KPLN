using System;
using System.Linq;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            // Material is an actual union of physical sections. Only its bounded holes
            // provide void evidence; neither a bounding box nor an open chain does so.
            internal static PlanarRegion EnclosedShaftVoids(PlanarRegion material)
            {
                if(material==null||material.IsEmpty)return PlanarRegion.Empty;
                var enclosed=PlanarRegion.Empty;
                foreach(var component in material.Components())
                    foreach(var hole in component.Skip(1))
                    {
                        var filled=PlanarRegion.FromRings(new[]{hole.Select(p=>new[]{p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale})});
                        enclosed=PlanarRegion.Combine(enclosed,filled,CT.ctUnion);
                    }
                // A hole can contain material islands, with further holes of their own.
                // Subtract all original material after the union to retain that topology.
                return PlanarRegion.Combine(enclosed,material,CT.ctDifference);
            }

            internal static bool ShaftMatchesPhysicalVoid(PlanarRegion candidate,PlanarRegion physicalVoid)
            {
                if(candidate==null||physicalVoid==null||candidate.IsEmpty||physicalVoid.IsEmpty)return false;
                if(PlanarRegion.Combine(candidate,physicalVoid,CT.ctXor).Area*.09290304>.001)return false;
                // Complete boundary coverage in both directions rejects a smaller loop
                // inside a larger void, equal-area translations and omitted material islands.
                // This only compensates integer-grid rounding, never a model displacement.
                return ShaftPerimeterMatches(candidate,physicalVoid)&&ShaftPerimeterMatches(physicalVoid,candidate);
            }

            internal static bool ShaftInteriorIsClear(PlanarRegion candidate,PlanarRegion material,bool complete)
            {
                if(!complete||candidate==null||candidate.IsEmpty||material==null)return false;
                // An opening in one slab does not prove empty space if another physical
                // section, such as a wall, occupies its interior at the same elevation.
                return PlanarRegion.Combine(candidate,material,CT.ctIntersection).Area*.09290304<=.001;
            }

            private static bool ShaftPerimeterMatches(PlanarRegion boundary,PlanarRegion support)
            {
                foreach(var ring in boundary.Paths)
                {
                    if(ring.Count<3)return false;
                    var points=ring.Select(p=>new[]{p.X/PlanarRegion.Scale,p.Y/PlanarRegion.Scale}).ToList();
                    points.Add(points[0]);
                    if(!PlanarRegion.SupportsBoundaryPolyline(PlanarRegion.Empty,new[]{support},points,4/PlanarRegion.Scale))return false;
                }
                return true;
            }
        }
    }
}
