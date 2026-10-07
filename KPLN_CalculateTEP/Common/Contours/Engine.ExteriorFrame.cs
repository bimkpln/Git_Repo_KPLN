using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private List<XYZ> ExteriorFramePoints(List<Record> walls,List<Record> physicalCandidates,double elevation)
            {
                var points=new List<XYZ>();var seen=new HashSet<string>();
                Action<Source,Element> add=(source,element)=>
                {
                    if(!seen.Add(source.Key+"/"+element.UniqueId))return;
                    var bounds=ElementBounds(element,source.Transform);
                    if(bounds!=null)
                    {
                        // The workspace includes elements touching the plane. A horizontal separator
                        // has no vertical extent and must not be rejected by half-open CrossesFloor.
                        if(elevation<bounds[2]-FloorGeometryTolerance||elevation>bounds[5]+FloorGeometryTolerance)return;
                        points.Add(new XYZ(bounds[0],bounds[1],elevation));
                        points.Add(new XYZ(bounds[3],bounds[4],elevation));
                        return;
                    }
                    var separator=element as ModelCurve;
                    if(separator==null)return;
                    var curve=separator.GeometryCurve;
                    if(curve==null)return;
                    var vertices=curve.Tessellate().Select(source.Transform.OfPoint).ToList();
                    if(vertices.Count==0||elevation<vertices.Min(p=>p.Z)-FloorGeometryTolerance||elevation>vertices.Max(p=>p.Z)+FloorGeometryTolerance)return;
                    points.AddRange(vertices.Select(p=>new XYZ(p.X,p.Y,elevation)));
                };
                foreach(var record in walls.Concat(physicalCandidates))add(record.Source,record.Element);
                foreach(var source in Sources.Where(s=>s.Loaded&&s.LoadError==null&&s.Mode=="include"))
                {
                    RequireHorizontalTransform(source.Transform);
                    foreach(var element in source.Elements.Where(e=>e is Wall||IsPhysicalShellCandidate(e)||
                        e is ModelCurve&&e.Category!=null&&IDHelper.ElIdValue(e.Category.Id)==(long)BuiltInCategory.OST_RoomSeparationLines))
                        if(PhaseAccepted(source,element))add(source,element);
                }
                return points;
            }

            private string DescribeExteriorFrameConflict(IList<BoundarySegment> ring,bool[] frame,double elevation)
            {
                string prefix="Граница временного внешнего помещения соединяет расчётную рамку с границей модели. Наружный контур этим способом не подтверждён. Отметка "+
                    (elevation*.3048).ToString("0.###",CultureInfo.InvariantCulture)+" м.";
                if(ring==null||ring.Count==0||frame==null||frame.Length!=ring.Count)return prefix;
                int index=-1;bool atStart=true;
                for(int i=0;i<ring.Count;i++)
                    if(!frame[i]&&frame[(i+ring.Count-1)%ring.Count]){index=i;break;}
                if(index<0)
                    for(int i=0;i<ring.Count;i++)
                        if(!frame[i]&&frame[(i+1)%ring.Count]){index=i;atStart=false;break;}
                if(index<0)index=Array.FindIndex(frame,value=>!value);
                if(index<0)return prefix;
                var segment=ring[index];
                var element=segment.ElementId==null||segment.ElementId==ElementId.InvalidElementId?null:doc.GetElement(segment.ElementId);
                string source=Sources.FirstOrDefault(s=>s.Document==doc)?.Name??doc.Title;
                string identity="ID границы "+(segment.ElementId==null?"не определён":IDHelper.ElIdValue(segment.ElementId).ToString(CultureInfo.InvariantCulture));
                var link=element as RevitLinkInstance;
                if(link!=null)
                {
                    source=Sources.FirstOrDefault(s=>s.RootLink==link.Id&&s.Document==link.GetLinkDocument())?.Name??link.Name;
                    identity+="; ID в связи "+(segment.LinkElementId==null?"не определён":IDHelper.ElIdValue(segment.LinkElementId).ToString(CultureInfo.InvariantCulture));
                    element=segment.LinkElementId==null||segment.LinkElementId==ElementId.InvalidElementId?null:link.GetLinkDocument()?.GetElement(segment.LinkElementId);
                }
                var point=segment.GetCurve().GetEndPoint(atStart?0:1);
                string kind=element is ModelCurve&&element.Category!=null&&IDHelper.ElIdValue(element.Category.Id)==(long)BuiltInCategory.OST_RoomSeparationLines?
                    "Разделитель помещений":element?.Category?.Name??"Элемент границы не определён";
                return prefix+" Ближайший к соединению участок: "+kind+"; источник «"+source+"»; "+identity+
                    "; XY "+(point.X*.3048).ToString("0.######",CultureInfo.InvariantCulture)+", "+(point.Y*.3048).ToString("0.######",CultureInfo.InvariantCulture)+" м.";
            }
        }
    }
}
