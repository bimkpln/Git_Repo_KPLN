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
            private static Schema StorageSchema()
            {
                var schema=Schema.Lookup(StorageId);if(schema!=null)return schema;
                var builder=new SchemaBuilder(StorageId);builder.SetSchemaName("KPLN_TEP_Spec20260826");
                builder.SetReadAccessLevel(AccessLevel.Public);builder.SetWriteAccessLevel(AccessLevel.Public);
                builder.AddSimpleField("Owner",typeof(string));builder.AddSimpleField("Kind",typeof(string));builder.AddSimpleField("Payload",typeof(string));return builder.Finish();
            }
            private static bool Owned(Element e)
            {
                var schema=Schema.Lookup(StorageId);if(schema==null)return false;
                var entity=e.GetEntity(schema);return entity.IsValid()&&entity.Get<string>(schema.GetField("Owner"))==Owner;
            }
            private static string Kind(Element e)
            {var s=Schema.Lookup(StorageId);if(s==null)return "";var entity=e.GetEntity(s);return entity.IsValid()?entity.Get<string>(s.GetField("Kind")):"";}
            private static void Tag(Element e,string kind,string payload="")
            {var s=StorageSchema();var entity=new Entity(s);entity.Set(s.GetField("Owner"),Owner);entity.Set(s.GetField("Kind"),kind);entity.Set(s.GetField("Payload"),payload??"");e.SetEntity(entity);}
            private T Read<T>(string kind) where T:class
            {
                var items=new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).Where(e=>Owned(e)&&Kind(e)==kind).OrderBy(e=>IDHelper.ElIdValue(e.Id)).ToList();
                if(kind=="settings"&&items.Count>1)throw new InvalidOperationException("В RVT несколько записей настроек ТЭП. Произвольная копия не выбрана; кнопка сохранения заменит их одной проверенной записью.");
                var item=items.FirstOrDefault();
                if(item==null)return null;var schema=StorageSchema();
                try{return Deserialize<T>(UnpackStorage(item.GetEntity(schema).Get<string>(schema.GetField("Payload"))));}
                catch(System.OperationCanceledException){throw;}
                    catch(Exception ex){throw new InvalidOperationException("Не удалось прочитать сохранённые данные ТЭП («"+kind+"»). Данные не перезаписаны. "+ex.Message,ex);}
            }
            private void Write<T>(string kind,T value)
            {
                string payload=PackStorage(Serialize(value));
                using(var t=new Transaction(doc,"ТЭП: сохранить "+kind))
                {
                    t.Start();var entries=new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).Where(e=>Owned(e)&&Kind(e)==kind).OrderBy(e=>IDHelper.ElIdValue(e.Id)).ToList();
                    var data=entries.FirstOrDefault() as DataStorage;
                    if(data==null){data=DataStorage.Create(doc);data.Name=Owner+"/"+kind;}
                    Tag(data,kind,payload);
                    if(kind=="settings")foreach(var duplicate in entries.Skip(1))doc.Delete(duplicate.Id);
                    if(t.Commit()!=TransactionStatus.Committed)throw new InvalidOperationException("Транзакция записи отменена.");
                }
            }
            private static string PackStorage(string text)
            {
                if(text.Length<1000000)return text;
                using(var output=new MemoryStream())
                {
                    using(var gzip=new System.IO.Compression.GZipStream(output,System.IO.Compression.CompressionMode.Compress,true))
                    {var bytes=Encoding.UTF8.GetBytes(text);gzip.Write(bytes,0,bytes.Length);}
                    string packed=CompressedStoragePrefix+Convert.ToBase64String(output.ToArray());
                    if(packed.Length>=16000000)throw new InvalidOperationException("Даже сжатый отчёт превышает лимит одного поля RVT. Экспортируйте отчёт и уменьшите набор показателей.");
                    return packed;
                }
            }
            private static string UnpackStorage(string text)
            {
                if(!text.StartsWith(CompressedStoragePrefix,StringComparison.Ordinal))return text;
                using(var input=new MemoryStream(Convert.FromBase64String(text.Substring(CompressedStoragePrefix.Length))))
                using(var gzip=new System.IO.Compression.GZipStream(input,System.IO.Compression.CompressionMode.Decompress))
                using(var reader=new StreamReader(gzip,Encoding.UTF8))return reader.ReadToEnd();
            }
            public static string Serialize<T>(T value)
            {using(var stream=new MemoryStream()){var serializer=new DataContractJsonSerializer(typeof(T),new DataContractJsonSerializerSettings{MaxItemsInObjectGraph=int.MaxValue});serializer.WriteObject(stream,value);return Encoding.UTF8.GetString(stream.ToArray());}}
            public static T Deserialize<T>(string text)
            {using(var stream=new MemoryStream(Encoding.UTF8.GetBytes(text))){return (T)new DataContractJsonSerializer(typeof(T),new DataContractJsonSerializerSettings{MaxItemsInObjectGraph=int.MaxValue}).ReadObject(stream);}}
            private sealed class Failures : IFailuresPreprocessor
            {
                private readonly Run run;public Failures(Run value){run=value;}
                public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
                {
                    bool errors=false;foreach(var failure in accessor.GetFailureMessages())
                    {
                        bool warning=failure.GetSeverity()==FailureSeverity.Warning;
                        run.Issue("REVIT_TRANSACTION",warning?"Предупреждение":"Ошибка",failure.GetDescriptionText(),element:string.Join(",",failure.GetFailingElementIds().Select(IDHelper.ElIdValue)));
                        if(warning)accessor.DeleteWarning(failure);else errors=true;
                    }
                    return errors?FailureProcessingResult.ProceedWithRollBack:FailureProcessingResult.Continue;
                }
            }
            private void Transaction(string name,Run run,Action action)
            {
                using(var t=new Transaction(doc,name))
                {
                    t.Start();t.SetFailureHandlingOptions(t.GetFailureHandlingOptions().SetFailuresPreprocessor(new Failures(run)).SetClearAfterRollback(true));
                    action();if(t.Commit()!=TransactionStatus.Committed)throw new InvalidOperationException("Revit отменил транзакцию «"+name+"».");
                }
            }
            private string UniqueName(string desired)
            {
                string safe=new string((desired??"ТЭП").Select(c=>"\\:{}[]|;<>?`~".Contains(c)?'_':c).ToArray());if(safe.Length>180)safe=safe.Substring(0,180);
                var names=new HashSet<string>(new FilteredElementCollector(doc).OfClass(typeof(View)).Select(v=>v.Name));string name=safe;int i=2;while(names.Contains(name))name=safe+" ("+(i++)+")";return name;
            }
        }
    }
}
