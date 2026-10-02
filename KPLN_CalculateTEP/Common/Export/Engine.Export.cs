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
            public static List<Tuple<string,List<object[]>>> ExportTables(Run run,int decimals)
            {
                var tables=new List<Tuple<string,List<object[]>>>();
                var totals=new List<object[]>{new object[]{"Код","Показатель","Исходное значение","Округлено","Ед.","Методика","Статус","Комментарий","Дата","Автор","Версия правил"}};
                totals.AddRange(run.Summary.Select(s=>new object[]{s.Key,s.Name,s.DisplayValue,s.NotCalculated?(object)null:Math.Round(s.Value,decimals,MidpointRounding.AwayFromZero),s.Unit,s.Method,s.Status,s.Comment,run.Date,run.Author,run.Version}));
                tables.Add(Tuple.Create("Итоги",totals));
                foreach(var kind in new[]{"Корпуса","Связи","Этажи"})
                {
                    var rows=new List<object[]>{new object[]{"Показатель","Группа","Секция","Этаж","Отметка, м","Исходное значение","Округлено","Ед.","Методика"}};
                    var groups=run.Details.Where(d=>!d.Excluded).GroupBy(d=>new{d.Metric,Group=kind=="Связи"?d.Source:d.Building,Section=kind=="Связи"?"":d.Section,Level=kind=="Этажи"?d.Level:"",Z=kind=="Этажи"?Math.Round(d.Elevation,6):0});
                    foreach(var g in groups)
                    {double value=g.Sum(d=>d.Value);if(kind=="Связи"&&(g.Key.Metric=="Storeys"||g.Key.Metric=="Floors"))value=g.GroupBy(d=>d.Building+"|"+d.Section).Select(x=>x.Sum(d=>d.Value)).DefaultIfEmpty(0).Max();
                        rows.Add(new object[]{Catalog().FirstOrDefault(m=>m.Key==g.Key.Metric)?.Name??g.Key.Metric,g.Key.Group,g.Key.Section,g.Key.Level,g.Key.Z*.3048,value,Math.Round(value,decimals,MidpointRounding.AwayFromZero),g.First().Unit,run.Method});}
                    tables.Add(Tuple.Create(kind,rows));
                }
                foreach(var name in new[]{"Детали площадей","Детали объёмов","Квартиры","Машино-места"})
                {
                    var rows=new List<object[]>{new object[]{"Показатель","Источник","Ключ источника","Корпус","Секция","Уровень","Отметка, м","ElementId","UniqueId","Назначение","ID квартиры","Профиль","Методика","Исходное значение","Коэффициент","Значение","Округлено","Ед.","Исключено","Вручную","Обоснование","Проверочный вид"}};
                    var selected=run.Details.Where(d=>name=="Квартиры"?d.Metric.StartsWith("Apartments"):name=="Машино-места"?d.Metric=="ParkingCount":name=="Детали объёмов"?d.Unit=="м³":d.Unit=="м²");
                    rows.AddRange(selected.Select(d=>new object[]{d.MetricName,d.Source,d.SourceKey,d.Building,d.Section,d.Level,d.Elevation*.3048,d.Element,d.UniqueId,d.PurposeName,d.Apartment,Choices("profile").FirstOrDefault(p=>p.Key==d.Profile)?.Label??d.Profile,d.Method,d.Raw,d.Factor,d.Value,Math.Round(d.Value,decimals,MidpointRounding.AwayFromZero),d.Unit,d.Excluded?"Да":"Нет",d.Manual?"Да":"Нет",d.Reason,d.ViewId}));
                    tables.Add(Tuple.Create(name,rows));
                }
                var issues=new List<object[]>{new object[]{"Код","Важность","Показатель","Источник","Корпус","Элемент","Сообщение","Что сделать"}};
                issues.AddRange(run.Issues.Select(x=>new object[]{x.Code,x.Severity,x.MetricName,x.Source,x.Building,x.Element,x.Message,x.Action}));tables.Add(Tuple.Create("Ошибки",issues));
                var settings=new List<object[]>{new object[]{"Раздел","Значение"},new object[]{"Методика",run.Method},new object[]{"Версия правил",run.Version},new object[]{"Дата",run.Date},new object[]{"Автор",run.Author}};
                string config=run.Configuration??"";for(int i=0;i<config.Length;i+=30000)settings.Add(new object[]{"Настройки JSON, часть "+(i/30000+1),config.Substring(i,Math.Min(30000,config.Length-i))});
                tables.Add(Tuple.Create("Методика и настройки",settings));
                var corrections=new List<object[]>{new object[]{"Показатель","Действие","Источник","Объект / корпус для дельты","Было","Новое значение / дельта","Причина","Автор","Дата"}};
                corrections.AddRange(run.Corrections.Select(c=>new object[]{c.Metric,c.Action,c.Source,c.Element,c.Previous,c.Value,c.Reason,c.Author,c.Date}));tables.Add(Tuple.Create("Ручные корректировки",corrections));
                return tables;
            }
            public static void ExportRun(Run run,string path,int decimals)
            {
                var tables=ExportTables(run,decimals);
                // Build to a sibling temporary file first: an interrupted export must not destroy an existing report.
                string temporary=path+"."+Guid.NewGuid().ToString("N")+".tmp";
                try
                {
                    if(Path.GetExtension(path).Equals(".xlsx",StringComparison.OrdinalIgnoreCase))WriteXlsx(temporary,tables);
                    else
                    {
                        using(var writer=new StreamWriter(temporary,false,new UTF8Encoding(true)))foreach(var table in tables)
                        {writer.WriteLine(CsvCell("РАЗДЕЛ: "+table.Item1));foreach(var row in table.Item2)writer.WriteLine(string.Join(";",row.Select(CsvCell)));writer.WriteLine();}
                    }
                    if(File.Exists(path))File.Replace(temporary,path,null);else File.Move(temporary,path);
                }
                finally{if(File.Exists(temporary))File.Delete(temporary);}
            }
            public static string CsvCell(object value)
            {
                string text=Convert.ToString(value,CultureInfo.InvariantCulture)??"";
                if(value is string&&text.TrimStart().Length>0&&"=+-@".Contains(text.TrimStart()[0]))text="'"+text;
                return "\""+text.Replace("\"","\"\"")+"\"";
            }
            private static string ColumnName(int index)
            {string text="";for(int n=index+1;n>0;n=(n-1)/26)text=(char)('A'+(n-1)%26)+text;return text;}
            private static string XmlText(object value)
            {var text=Convert.ToString(value,CultureInfo.InvariantCulture)??"";return new string(text.Where(c=>XmlConvert.IsXmlChar(c)).ToArray());}
            private static void WriteXlsx(string path,List<Tuple<string,List<object[]>>> tables)
            {
                XNamespace ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main",rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                using(var package=Package.Open(path,FileMode.Create,FileAccess.ReadWrite))
                {
                    var workbookUri=new Uri("/xl/workbook.xml",UriKind.Relative);
                    var workbook=package.CreatePart(workbookUri,"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",CompressionOption.Normal);
                    package.CreateRelationship(workbookUri,TargetMode.Internal,rel.NamespaceName+"/officeDocument");
                    var sheets=new XElement(ns+"sheets");int sheetIndex=0;
                    foreach(var table in tables)
                    {
                        sheetIndex++;if(table.Item2.Count>1048576)throw new InvalidOperationException("Лист превышает лимит Excel. Экспортируйте CSV.");
                        var uri=new Uri("/xl/worksheets/sheet"+sheetIndex+".xml",UriKind.Relative);
                        var sheet=package.CreatePart(uri,"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",CompressionOption.Normal);
                        string rid="rId"+sheetIndex;workbook.CreateRelationship(PackUriHelper.GetRelativeUri(workbookUri,uri),TargetMode.Internal,rel.NamespaceName+"/worksheet",rid);
                        sheets.Add(new XElement(ns+"sheet",new XAttribute("name",table.Item1),new XAttribute("sheetId",sheetIndex),new XAttribute(rel+"id",rid)));
                        var data=new XElement(ns+"sheetData");int rowIndex=0;
                        foreach(var values in table.Item2)
                        {
                            rowIndex++;var row=new XElement(ns+"row",new XAttribute("r",rowIndex));int col=0;
                            foreach(var value in values)
                            {
                                var cell=new XElement(ns+"c",new XAttribute("r",ColumnName(col++)+rowIndex));
                                if(value is double||value is int||value is long||value is decimal)cell.Add(new XElement(ns+"v",Convert.ToString(value,CultureInfo.InvariantCulture)));
                                else
                                {
                                    var text=XmlText(value);if(text.Length>32767)throw new InvalidOperationException("Текст ячейки превышает лимит Excel. Экспортируйте CSV.");
                                    cell.Add(new XAttribute("t","inlineStr"),new XElement(ns+"is",new XElement(ns+"t",new XAttribute(XNamespace.Xml+"space","preserve"),text)));
                                }
                                row.Add(cell);
                            }
                            data.Add(row);
                        }
                        var root=new XElement(ns+"worksheet",new XElement(ns+"sheetViews",new XElement(ns+"sheetView",new XAttribute("workbookViewId",0),new XElement(ns+"pane",new XAttribute("ySplit",1),new XAttribute("topLeftCell","A2"),new XAttribute("activePane","bottomLeft"),new XAttribute("state","frozen")))),
                            new XElement(ns+"cols",new XElement(ns+"col",new XAttribute("min",1),new XAttribute("max",table.Item2[0].Length),new XAttribute("width",24),new XAttribute("customWidth",1))),data,
                            new XElement(ns+"autoFilter",new XAttribute("ref","A1:"+ColumnName(table.Item2[0].Length-1)+Math.Max(1,rowIndex))));
                        using(var stream=sheet.GetStream())new XDocument(root).Save(stream);
                    }
                    using(var stream=workbook.GetStream())new XDocument(new XElement(ns+"workbook",new XAttribute(XNamespace.Xmlns+"r",rel),sheets)).Save(stream);
                }
            }
        }
    }
}
