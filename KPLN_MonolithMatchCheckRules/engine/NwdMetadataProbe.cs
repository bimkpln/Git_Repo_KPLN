using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Security.Cryptography;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Controls;

class NwdMetadataProbe {
 static readonly string Runtime=Environment.GetEnvironmentVariable("NAVISWORKS_RUNTIME")??@"C:\Program Files\Autodesk\Navisworks Manage 2023";
 static string Hash(string path){using(var stream=File.OpenRead(path))using(var sha=SHA256.Create())return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","").ToLowerInvariant();}
 [STAThread]static int Main(string[] args){
  AppDomain.CurrentDomain.AssemblyResolve+=(s,e)=>{string p=Path.Combine(Runtime,new AssemblyName(e.Name).Name+".dll");return File.Exists(p)?Assembly.LoadFrom(p):null;};
  Environment.SetEnvironmentVariable("PATH",Runtime+";"+Environment.GetEnvironmentVariable("PATH"));
  try{
   if(args.Length!=2)throw new ArgumentException("input.nwd output-directory required");
   string input=Path.GetFullPath(args[0]),output=Path.GetFullPath(args[1]);
   if(!input.EndsWith(".nwd",StringComparison.OrdinalIgnoreCase))throw new ArgumentException("NWD only");
   Directory.CreateDirectory(output);
   if(File.Exists(Path.Combine(output,"nodes.json"))||File.Exists(Path.Combine(output,"manifest.json")))throw new IOException("Existing metadata is immutable; use a new output directory");
   return Read(input,output);
  }catch(Exception e){Console.Error.WriteLine(e);return 1;}
 }
 static void Hide(SavedItem item,bool hidden,Document doc,HashSet<ModelItem> excluded){
  hidden=hidden||item.DisplayName.Trim()=="Скрыть";
  var group=item as GroupItem;if(group!=null)foreach(var child in group.Children)Hide(child,hidden,doc,excluded);
  var set=item as SelectionSet;if(set!=null&&hidden)foreach(var selected in set.GetSelectedItems(doc))foreach(var node in selected.DescendantsAndSelf)excluded.Add(node);
 }
 [MethodImpl(MethodImplOptions.NoInlining)]static int Read(string input,string output){
  string before=Hash(input);var watch=System.Diagnostics.Stopwatch.StartNew();
  ApplicationControl.ApplicationType=ApplicationType.SingleDocument;ApplicationControl.Initialize();
  try{using(var control=new DocumentControl()){
   control.SetAsMainDocument();if(!control.Document.TryOpenFile(input))throw new Exception("NWD could not be opened");
   var excluded=new HashSet<ModelItem>();Hide(control.Document.SelectionSets.RootItem,false,control.Document,excluded);
   var serializer=new JavaScriptSerializer();serializer.MaxJsonLength=int.MaxValue;
   var indices=new Dictionary<ModelItem,int>();var sourceNames=new Dictionary<ModelItem,string>();var models=new List<object>();int count=0;
   using(var writer=new StreamWriter(new FileStream(Path.Combine(output,"nodes.json"),FileMode.CreateNew),new System.Text.UTF8Encoding(false))){
    writer.Write("[");bool first=true;
    foreach(var model in control.Document.Models){
     models.Add(new {file=model.FileName,units=model.Units.ToString()});
     foreach(var item in model.RootItem.DescendantsAndSelf){
      int parent=-1;if(item.Parent!=null&&!indices.TryGetValue(item.Parent,out parent))throw new Exception("Parent missing from traversal");
      string sourceName=item.DisplayName.EndsWith(".nwc",StringComparison.OrdinalIgnoreCase)?item.DisplayName:item.Parent==null?model.FileName:sourceNames[item.Parent];
      sourceNames.Add(item,sourceName);
      var props=new List<object>();foreach(var category in item.PropertyCategories)foreach(var p in category.Properties)
       props.Add(new {category=category.DisplayName,category_name=category.Name,name=p.DisplayName,property_name=p.Name,value=p.Value.ToString()});
      double[] bounds=null;if(props.Count>0){var b=item.BoundingBox();if(!b.IsEmpty)bounds=new[]{b.Min.X,b.Min.Y,b.Min.Z,b.Max.X,b.Max.Y,b.Max.Z};}
      var record=new {index=count,parent=parent,model=model.FileName,source=sourceName,name=item.DisplayName,has_geometry=item.HasGeometry,
       excluded_by_hide_set=excluded.Contains(item),hidden=item.IsHidden,bounds=bounds,properties=props};
      if(!first)writer.Write(",");first=false;writer.Write(serializer.Serialize(record));indices.Add(item,count++);
      if(count%10000==0)Console.WriteLine("METADATA NODES "+count);
     }
    }writer.Write("]");
   }
   string after=Hash(input);if(before!=after)throw new Exception("Input NWD changed during metadata read");
   File.WriteAllText(Path.Combine(output,"manifest.json"),serializer.Serialize(new {schema=1,complete=true,input=input,input_sha256=before,input_sha256_after_read=after,models=models,
    nodes=count,excluded_set_nodes=excluded.Count,seconds=watch.Elapsed.TotalSeconds,source_modified=false,geometry_extracted=false,
    purpose="Metadata only; no geometric comparison or complete-check acceptance",reader_executable_sha256=Hash(Assembly.GetExecutingAssembly().Location),metadata_sha256=Hash(Path.Combine(output,"nodes.json"))}));
   Console.WriteLine("METADATA COMPLETE "+count);return 0;
  }}finally{ApplicationControl.Terminate();}
 }
}
