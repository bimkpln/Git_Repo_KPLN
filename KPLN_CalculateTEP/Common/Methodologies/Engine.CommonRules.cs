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
            private static bool IsPublic(Record r){return r.Profile=="public"||r.Profile=="high-public";}
            private static bool IsHigh(Record r){return r.Profile.StartsWith("high-");}
            private static bool OneOf(string value,params string[] values){return values.Contains(value);}
            private static bool GrossMetric(Indicator metric)
            {return OneOf(metric.ToString(),"Gross","GrossAbove","GrossBelow","Np","NpResidential","NpNonresidential","NnpEmbedded","NnpSeparate");}
            private bool Above(Record r,bool forStoreys=false)
            {
                if(r.Level==null)throw new InvalidOperationException("Не определён расчётный этаж объекта.");
                if(r.Level.Above=="above")return true;if(r.Level.Above=="below")return false;
                if(r.Level.Kind=="basement")
                {
                    if(forStoreys&&IsPublic(r)&&!IsHigh(r))return true;
                    double top=ResolvedTopSlab(r.Level);
                    return top-ground>=2/.3048-1e-8;
                }
                return r.Z>=ground-1e-6;
            }
            private bool CountLevel(Record r,bool aboveOnly)
            {
                if(r.Level==null||!r.Level.Include||r.Level.Kind=="exclude")return false;
                if(r.Level.Kind=="attic")return false;
                if(r.Level.Kind=="roof"||OneOf(r.Role,"roof-exit","roof-vent"))
                {
                    if(IsHigh(r))
                    {
                        double area=ResolvedLevelField(r.Level,"area");
                        double height=ResolvedLevelField(r.Level,"height");
                        if(area<8&&height<2.5)return false;
                    }
                    else if(!IsPublic(r)||r.Role=="roof-exit")return false;
                    else if(ResolvedLevelField(r.Level,"ratio")<.15)return false;
                }
                if(r.Level.Kind=="void")
                {
                    if(!IsPublic(r)||IsHigh(r))return false;
                    return ResolvedLevelField(r.Level,"height")>=1.8;
                }
                if(IsHigh(r)&&OneOf(r.Role,"technical-space","technical-void")&&AutomaticDimension(r,"partial-floor")>0)return false;
                return !aboveOnly||Above(r,true);
            }
            private string Part(Record r)
            {
                if(r.Part!="auto")return r.Part;
                if(OneOf(r.Role,"heated","auxiliary","residential-common","entrance")||!string.IsNullOrWhiteSpace(r.Apartment))return "residential";
                if(OneOf(r.Role,"public","nonresidential-common","parking"))return "nonresidential";
                throw new InvalidOperationException("Для роли «"+r.Role+"» задайте обслуживаемую жилую / нежилую часть.");
            }
            private bool Eligible(Record r,Indicator metric)
            {
                if(r.Override.HasValue)return r.Override.Value;
                if(r.Level==null||!r.Level.Include||r.Level.Kind=="exclude")return false;
                if(r.Role=="unknown")throw new InvalidOperationException("Назначение объекта не определено правилами или параметром функции.");
                string role=r.Role;bool pub=IsPublic(r),gns=(int)metric<=4;
                if(gns)
                {
                    if(!Above(r))return false;
                    if(metric==Indicator.GnsResidential&&r.BuildingClass!="residential")return false;
                    if(metric==Indicator.GnsNonresidential&&r.BuildingClass!="nonresidential")return false;
                    if(metric==Indicator.GnsLivingPart&&(r.BuildingClass!="residential"||Part(r)!="residential"))return false;
                    if(metric==Indicator.GnsNonlivingPart&&(r.BuildingClass!="residential"||Part(r)!="nonresidential"))return false;
                    if(OneOf(role,"technical-void","attic","roof-exit","roof-vent","open-stair","ramp","canopy","porch","pit","passage","decoration","french-balcony","apartment-family","ventilated-void","soil-filled"))return false;
                    if(Methodology.ExcludesGnsRole(role))return false;
                    return !OneOf(role,"structure","envelope","footprint","underground-footprint");
                }
                if((metric==Indicator.GrossAbove||OneOf(metric.ToString(),"Np","NpResidential","NpNonresidential","NnpEmbedded","NnpSeparate"))&&!Above(r))return false;
                if(metric==Indicator.GrossBelow&&Above(r))return false;
                if(metric==Indicator.NpResidential&&r.BuildingClass!="residential")return false;
                if(metric==Indicator.NpNonresidential&&r.BuildingClass!="nonresidential")return false;
                if(metric==Indicator.NnpEmbedded||metric==Indicator.NnpSeparate)
                {
                    if(Part(r)!="nonresidential")return false;
                    if(string.IsNullOrWhiteSpace(Config.Parameter("embedded"))||string.IsNullOrWhiteSpace(Config.Parameter("standalone")))
                        throw new InvalidOperationException("Для ННП задайте признаки встроенно-пристроенной части и отдельно стоящего объекта.");
                    if(string.IsNullOrWhiteSpace(Mapped(r.Element,"embedded"))||string.IsNullOrWhiteSpace(Mapped(r.Element,"standalone")))throw new InvalidOperationException("У нежилого помещения не заполнены признаки встроенно-пристроенной / отдельно стоящей части.");
                    if(!Flag(r.Element,"embedded"))return false;
                    if(Flag(r.Element,"standalone")!=(metric==Indicator.NnpSeparate))return false;
                }
                if(GrossMetric(metric))
                {
                    if(OneOf(role,"structure","envelope","footprint","underground-footprint","decoration","stove","canopy","porch","pit","passage","french-balcony","open-stair","apartment-family"))return false;
                    if(!pub&&role=="ramp")return false;
                    if(!pub&&OneOf(role,"technical-void","technical-space","attic","roof-exit","roof-vent","external-tambour","tambour","engineering-shaft"))return false;
                    if(pub&&OneOf(role,"attic","roof-exit"))return false;
                    if(pub&&role=="roof-vent")return Dimension(r,"roof-ratio")>=.15;
                    if(OneOf(role,"ventilated-void","soil-filled"))return false;
                    if(pub&&OneOf(role,"technical-void","technical-space")&&Dimension(r,"height",true)<1.8)
                    {if(string.IsNullOrWhiteSpace(Mapped(r.Element,"service-access")))throw new InvalidOperationException("Для низкого технического пространства укажите, нужен ли проход обслуживания коммуникаций.");return Flag(r.Element,"service-access");}
                    if(IsHigh(r)&&OneOf(role,"technical-space","technical-void")&&AutomaticDimension(r,"partial-floor")>0)return false;
                    if(pub&&role=="mezzanine")return r.SingleStorey||Dimension(r,"mezzanine-ratio")>.4;
                    return true;
                }
                if(metric==Indicator.ApartmentsTotal||metric==Indicator.ApartmentsHeated)
                {
                    if(string.IsNullOrWhiteSpace(r.Apartment))return false;
                    if(metric==Indicator.ApartmentsHeated)return OneOf(role,"heated","auxiliary","under-stair","niche","arch","mezzanine");
                    return OneOf(role,"heated","auxiliary","under-stair","niche","arch","mezzanine","balcony","loggia","terrace","veranda","cold-storage","tambour");
                }
                if(metric==Indicator.PublicRooms)return role=="public"||r.Part=="nonresidential"&&OneOf(role,"external-tambour","balcony","loggia","terrace","veranda","stair","multilight","opening","shaft");
                if(metric==Indicator.PublicCalculated)
                    return (pub||role=="public"||r.Part=="nonresidential")&&OneOf(role,"public","heated","auxiliary","niche","arch","mezzanine");
                return false;
            }
            private double Dimension(Record r,string key,bool length=false)
            {if(OneOf(key,"width","stair-width","roof-ratio","mezzanine-ratio","height"))return AutomaticDimension(r,key);var n=Number(r.Element,key,length);if(!n.HasValue)throw new InvalidOperationException("Нужен параметр: "+Config.Parameters.First(x=>x.Key==key).Title);return n.Value;}
            private double Factor(Record r,Indicator metric)
            {
                if(metric!=Indicator.ApartmentsTotal)return 1;
                double factor=OneOf(r.Role,"balcony","terrace")?SummerCoefficient(Config.BalconyCoefficient):r.Role=="loggia"?SummerCoefficient(Config.LoggiaCoefficient):1;
                if(factor<0||factor>1)throw new InvalidOperationException("Коэффициент должен быть в диапазоне 0..1.");return factor;
            }
            public static double PublicMansardHeight(double angle)
            {if(angle<0||angle>90)throw new ArgumentOutOfRangeException("angle");return angle<=30?1.5:angle<=45?1.5-(angle-30)*.4/15:angle<=60?1.1-(angle-45)*.6/15:.5;}
            private bool VerticalExclusion(Record r,Indicator metric,List<Record> all,Solid shape)
            {return VerticalExclusionArea(r,metric,all,shape?.Volume??0);}
            private bool VerticalExclusionArea(Record r,Indicator metric,List<Record> all,double areaFeet)
            {
                bool gns=(int)metric<=4;
                bool stair=r.Role=="stair-gap"&&(Dimension(r,"width",true)>1.5||Dimension(r,"width",true)>Dimension(r,"stair-width",true));
                bool vertical=r.Role=="multilight"||stair||r.Role=="opening"&&(!gns||Methodology.CountsGnsOpeningAreaOnOneFloor(areaFeet*.09290304))||OneOf(r.Role,"shaft","engineering-shaft")&&(!gns||Methodology.CountsGnsShaftOnOneFloor);
                if(!vertical)return false;
                if(string.IsNullOrWhiteSpace(r.Vertical))
                {
                    if(r.Element is Opening)
                    {Notice("VERTICAL_LEVELS","Ошибка","Проём задан одним объектом без поэтажных контуров. Для вычитания на верхних этажах задайте зоны с общим ID вертикального пространства.",r,metric.ToString());return false;}
                    throw new InvalidOperationException("Для многосветного пространства / проёма / шахты нужен общий ID вертикального пространства на всех этажах.");
                }
                var matching=all.Where(x=>x.Building==r.Building&&x.Section==r.Section&&x.Vertical==r.Vertical&&x.Level!=null&&x.Level.Include).ToList();
                if(matching.Count==0)return false;return r.Z>matching.Min(x=>x.Z)+1e-6;
            }
        }
    }
}
