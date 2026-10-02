using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using KPLN_CalculateTEP.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using PC = KPLN_CalculateTEP.Common.Geometry.TepClipper.Clipper;
using CP = KPLN_CalculateTEP.Common.Geometry.TepClipper.IntPoint;
using CT = KPLN_CalculateTEP.Common.Geometry.TepClipper.ClipType;
using PT = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyType;
using PF = KPLN_CalculateTEP.Common.Geometry.TepClipper.PolyFillType;

using TepClipper = KPLN_CalculateTEP.Common.Geometry.TepClipper;
using KPLN_CalculateTEP.Common.Methodologies;
namespace KPLN_CalculateTEP.Common
{
    public static partial class TepCalculation
    {
        public partial class Engine
        {
            private Solid SpatialPlan(Record r,string boundary,string metric="")
            {
                if(r.Element is Room)return RoomPlanAtFloor(r,boundary,metric);
                var spatial=(SpatialElement)r.Element;
                var options=new SpatialElementBoundaryOptions{SpatialElementBoundaryLocation=boundary=="net"?SpatialElementBoundaryLocation.Finish:SpatialElementBoundaryLocation.Center};
                var boundaries=spatial.GetBoundarySegments(options);
                if(boundaries==null||boundaries.Count==0)throw new InvalidOperationException("Нет замкнутых границ помещения / зоны.");
                var loops=boundaries.Select(b=>FlatLoop(b.Select(x=>x.GetCurve()))).ToList();
                if(!(spatial is Area)&&boundary!="net")
                {
                    var centre=Extrude(loops);var modified=new List<CurveLoop>();
                    for(int li=0;li<loops.Count;li++)
                    {
                        var distances=new List<double>();bool changed=false;
                        foreach(var segment in boundaries[li])
                        {
                            var boundaryElement=r.Element.Document.GetElement(segment.ElementId);var wall=boundaryElement as Wall;var wallTransform=Transform.Identity;double d=0;
                            var boundaryLink=boundaryElement as RevitLinkInstance;
                            if(boundaryLink!=null){wall=boundaryLink.GetLinkDocument()?.GetElement(segment.LinkElementId) as Wall;wallTransform=boundaryLink.GetTotalTransform();if(wall==null)throw new InvalidOperationException("Граница связана со связью, но исходная стена недоступна. Задайте расчётный контур зоной.");}
                            if(wall!=null&&wall.WallType.Function==WallFunction.Exterior)
                            {
                                if(wall.WallType.Kind!=WallKind.Basic)throw new InvalidOperationException("Нестандартная наружная стена: задайте внешний / внутренний контур зоной.");
                                BuiltInParameter sectionParameter;if(Enum.TryParse("WALL_CROSS_SECTION",out sectionParameter))
                                {var crossSection=wall.get_Parameter(sectionParameter);if(crossSection!=null&&crossSection.AsInteger()!=0)throw new InvalidOperationException("Наклонная / коническая наружная стена требует явного контура на нормативной высоте обмера.");}
                                var c=Flat(segment.GetCurve());var p=c.Evaluate(.5,true);var tangent=c.ComputeDerivatives(.5,true).BasisX.Normalize();
                                var right=tangent.CrossProduct(XYZ.BasisZ).Normalize();double sign=Inside(centre,p+right*.002)?-1:1;
                                double width=wall.Width/2;
                                if(boundary=="outer"&&Methodology.UsesCoreBoundary)
                                {var compound=wall.WallType.GetCompoundStructure();if(compound!=null){var layers=compound.GetLayers();int first=compound.GetFirstCoreLayerIndex();if(first<0)throw new InvalidOperationException("У наружной стены не задан несущий слой для нового ГНС.");
                                    double finish=0;bool exterior=wallTransform.OfVector(wall.Orientation).DotProduct(right*sign)>0;
                                    if(exterior){for(int i=0;i<first;i++)finish+=layers[i].Width;}
                                    else{for(int i=compound.GetLastCoreLayerIndex()+1;i<layers.Count;i++)finish+=layers[i].Width;}
                                    width-=finish;}}
                                d=sign*(boundary=="outer"?width:-wall.Width/2);changed|=Math.Abs(d)>1e-9;
                            }
                            else if(segment.ElementId==ElementId.InvalidElementId&&boundary=="outer")
                                Notice("BOUNDARY_SEPARATOR","Предупреждение","Часть границы не связана с наружной стеной. Проверьте внешний контур на расчётном виде.",r,metric);
                            distances.Add(d);
                        }
                        modified.Add(changed?CurveLoop.CreateViaOffset(loops[li],distances,XYZ.BasisZ):loops[li]);
                    }
                    loops=modified;
                }
                var local=Extrude(loops);
                var horizontal=Transform.CreateTranslation(new XYZ(0,0,-r.Source.Transform.Origin.Z)).Multiply(r.Source.Transform);
                return SolidUtils.CreateTransformed(local,horizontal);
            }
            private Solid Plan(Record r,Indicator metric)
            {
                Progress("Геометрия: "+MetricLabels[metric.ToString()]+" / "+r.Source.Name+" / "+r.Level?.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                string boundary=(int)metric<=4?"outer":GrossMetric(metric)?"gross":"net";
                if(boundary=="outer"&&r.Element is SpatialElement&&!(r.Element is Area))Notice("EXTERIOR_FUNCTION","Предупреждение","Автоматический внешний контур использует функцию наружных стен и их ядро. Проверьте маркировку наружных стен и полноту помещений на служебном плане.",r,metric.ToString());
                if(r.Element is Area)Notice("AREA_CONTOUR","Предупреждение","Используется готовая граница зоны. Проверьте соответствие схемы: наружный контур для ГНС, внутренний для общей / наземной площади, чистовой для квартир и помещений.",r,metric.ToString());
                if(boundary=="net"&&!(r.Element is SpatialElement)&&!(r.Element is Opening))Notice("FAMILY_NET_CONTOUR","Предупреждение","Для семейства используется горизонтальная проекция. Проверьте чистовой контур и исключение толщины ограждений; при необходимости используйте отдельную зону.",r,metric.ToString());
                string key=r.Element is Room?RoomPlanCacheKey(r,boundary,metric.ToString()):r.Key+"/"+boundary;Solid shape;
                if(!shapes.TryGetValue(key,out shape))
                {
                    if(r.Element is SpatialElement)shape=SpatialPlan(r,boundary,metric.ToString());
                    else if(r.Element is Opening)
                    {
                        var opening=(Opening)r.Element;var curves=opening.BoundaryCurves.Cast<Curve>().ToList();
                        if(curves.Count==0)throw new InvalidOperationException("У проёма нет горизонтальной границы. Назначьте расчётную зону.");
                        var transformed=curves.Select(c=>c.CreateTransformed(r.Source.Transform));shape=Extrude(new List<CurveLoop>{FlatLoop(transformed)});
                    }
                    else
                    {
                        shape=null;foreach(var solid in Solids(r.Element))shape=Union(shape,Projection(SolidUtils.CreateTransformed(solid,r.Source.Transform)));
                        if(shape==null)throw new InvalidOperationException("Нет замкнутой геометрии объекта.");
                    }
                    shapes[key]=shape;
                }
                if(boundary!="outer")
                {
                    double min=0;
                    bool residentialRoom=metric==Indicator.ApartmentsTotal||metric==Indicator.ApartmentsHeated||!IsPublic(r)&&r.Part!="nonresidential"&&r.Role!="public";
                    if(residentialRoom&&r.Role=="under-stair")min=1.600001;
                    if(residentialRoom&&r.Role=="niche"&&Dimension(r,"height",true)<2)return null;
                    if(residentialRoom&&r.Role=="arch"&&Dimension(r,"width",true)<2)return null;
                    var slope=Number(r.Element,"slope");
                    if(slope.HasValue)
                    {
                        if(!IsPublic(r)&&r.Role!="public"&&r.Part!="nonresidential")
                        {if(r.Element is Area&&r.Override==true){Notice("MANSARD_MANUAL","Предупреждение","Использован явно включённый контур жилой мансарды. Проверьте согласованное обоснование высотной границы.",r,metric.ToString());return shape;}
                            throw new InvalidOperationException("Для жилой мансарды в исходном ТЗ неоднозначно заданы пороги высоты. Подготовьте зону учитываемой части и отдельное правило включения с обоснованием.");}
                        min=metric==Indicator.PublicCalculated?1.5:PublicMansardHeight(slope.Value);
                    }
                    if(min>0&&r.Element is SpatialElement&&!(r.Element is Area))
                    {
                        var volume=SpatialVolume(r,boundary=="gross");var section=Section(volume,(r.Element is Room?RoomFloorElevation(r):r.Z)+min/.3048);
                        shape=Intersect(shape,section);
                    }
                    else if(min>0&&r.Element is Area)
                        Notice("AREA_HEIGHT","Предупреждение","Зона должна быть заранее обрезана по нормативной высоте: "+min.ToString("0.###")+" м.",r,metric.ToString());
                }
                return shape;
            }
            private Solid SpatialVolume(Record r,bool centre=false)
            {
                string key=r.Key+(centre?"/volume-centre":"/volume");Solid cached;if(shapes.TryGetValue(key,out cached))return cached;
                if(r.Element is Area)throw new InvalidOperationException("Зона не содержит высоту и не определяет строительный объём. Нужны Rooms/Spaces и оболочка либо замкнутый расчётный объём.");
                var localKey=Tuple.Create(r.Element,centre);Solid local;
                if(!localSpatialVolumes.TryGetValue(localKey,out local))
                {
                    Progress("Геометрия помещения / пространства: "+r.Source.Name+"; ID "+IDHelper.ElIdValue(r.Element.Id));
                    local=Measure("Геометрия Rooms/Spaces",r,"",()=>
                    {
                        var calculatorKey=Tuple.Create(r.Element.Document,centre);SpatialElementGeometryCalculator calculator;
                        if(!spatialCalculators.TryGetValue(calculatorKey,out calculator))
                            spatialCalculators[calculatorKey]=calculator=new SpatialElementGeometryCalculator(r.Element.Document,new SpatialElementBoundaryOptions{SpatialElementBoundaryLocation=centre?SpatialElementBoundaryLocation.Center:SpatialElementBoundaryLocation.Finish});
                        return calculator.CalculateSpatialElementGeometry((SpatialElement)r.Element).GetGeometry();
                    });
                    localSpatialVolumes.Add(localKey,local);
                }
                else spatialCacheHits++;
                cached=SolidUtils.CreateTransformed(local,r.Source.Transform);shapes[key]=cached;return cached;
            }
        }
    }
}
