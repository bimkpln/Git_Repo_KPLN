using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Controls;
using Autodesk.Navisworks.Api.ComApi;
using Com=Autodesk.Navisworks.Api.Interop.ComApi;

[ComVisible(true)] public class Triangles : Com.InwSimplePrimitivesCB {
  public BinaryWriter Writer;
  public BinaryWriter LocalWriter;
  public int Count;
  public double[] Matrix;
  public void Triangle(Com.InwSimpleVertex a, Com.InwSimpleVertex b, Com.InwSimpleVertex c) { Vertex(a);Vertex(b);Vertex(c);Count++; }
  void Vertex(Com.InwSimpleVertex v) {
    var a=(Array)v.coord; int k=a.GetLowerBound(0);
    double x=Convert.ToDouble(a.GetValue(k)),y=Convert.ToDouble(a.GetValue(k+1)),z=Convert.ToDouble(a.GetValue(k+2));
    LocalWriter.Write(x);LocalWriter.Write(y);LocalWriter.Write(z);
    Writer.Write(x*Matrix[0]+y*Matrix[4]+z*Matrix[8]+Matrix[12]);
    Writer.Write(x*Matrix[1]+y*Matrix[5]+z*Matrix[9]+Matrix[13]);
    Writer.Write(x*Matrix[2]+y*Matrix[6]+z*Matrix[10]+Matrix[14]);
  }
  public void Line(Com.InwSimpleVertex a,Com.InwSimpleVertex b) {}
  public void Point(Com.InwSimpleVertex a) {}
  public void SnapPoint(Com.InwSimpleVertex a) {}
}

class NwdFragmentProbe {
  static readonly string Runtime = Environment.GetEnvironmentVariable("NAVISWORKS_RUNTIME") ?? @"C:\Program Files\Autodesk\Navisworks Manage 2023";
  [STAThread] static int Main(string[] args) {
    AppDomain.CurrentDomain.AssemblyResolve += (s,e) => {
      string p=Path.Combine(Runtime,new AssemblyName(e.Name).Name+".dll");
      return File.Exists(p) ? Assembly.LoadFrom(p) : null;
    };
    Environment.SetEnvironmentVariable("PATH",Runtime+";"+Environment.GetEnvironmentVariable("PATH"));
    try {
      if(args.Length!=2)throw new ArgumentException("Usage: NwdFragmentProbe.exe input.nwd new-output-directory");
      string input=Path.GetFullPath(args[0]);
      if(!File.Exists(input))throw new FileNotFoundException("Input NWD unavailable",input);
      string output=Path.GetFullPath(args[1]);
      if(Directory.Exists(output))throw new IOException("Output directory already exists; choose a new one");
      Directory.CreateDirectory(output);Directory.SetCurrentDirectory(output);
      File.WriteAllText("export_started.json","{\"complete\":false,\"schema\":2}");
      return Probe(input);
    } catch(Exception e) { Console.WriteLine(e); return 1; }
  }
  static string FileHash(string file) {
    using(var h=System.Security.Cryptography.SHA256.Create())using(var f=File.OpenRead(file))
      return BitConverter.ToString(h.ComputeHash(f)).Replace("-","").ToLowerInvariant();
  }
  static void Collect(SavedItem item, bool inHide, string parent, Document doc, HashSet<ModelItem> excluded, List<object> sets) {
    string path=parent+"/"+item.DisplayName;
    bool hide=inHide || item.DisplayName.Trim()==(Environment.GetEnvironmentVariable("NAVIS_HIDE_FOLDER") ?? "Скрыть");
    var group=item as GroupItem;
    if(group!=null) foreach(var child in group.Children) Collect(child,hide,path,doc,excluded,sets);
    var set=item as SelectionSet;
    if(set!=null) {
      var selected=set.GetSelectedItems(doc);
      sets.Add(new {path=path,excluded=hide,count=selected.Count});
      if(hide) foreach(var model in selected) foreach(var node in model.DescendantsAndSelf) excluded.Add(node);
    }
  }
  [MethodImpl(MethodImplOptions.NoInlining)] static int Probe(string file) {
    var watch=System.Diagnostics.Stopwatch.StartNew();double readSeconds=0;
    Console.WriteLine("Initializing Navisworks reader");
    ApplicationControl.ApplicationType=ApplicationType.SingleDocument;
    ApplicationControl.Initialize();
    try {
      using (var control=new DocumentControl()) {
        control.SetAsMainDocument();
        Console.WriteLine("Loading NWD");
        if(!control.Document.TryOpenFile(file)) throw new Exception("NWD reader could not open file");
        readSeconds=watch.Elapsed.TotalSeconds;
        var excluded = new HashSet<ModelItem>();
        var sets = new List<object>();
        Collect(control.Document.SelectionSets.RootItem, false, "", control.Document, excluded, sets);
        File.WriteAllText("nwd_sets.json",new JavaScriptSerializer().Serialize(sets));
        Console.WriteLine("Excluded set nodes: "+excluded.Count);
        int count=0,geometry=0;
        var json=new JavaScriptSerializer(); json.MaxJsonLength=int.MaxValue;
        var propertiesCache=new Dictionary<ModelItem,List<object>>();
        var models=new List<object>();var scene=new List<object>();var sceneIndices=new Dictionary<ModelItem,int>();
        using(var metadata=new StreamWriter("nwd_geometry.json",false,new System.Text.UTF8Encoding(false))) {
        metadata.Write("[");bool first=true;
        using(var writer=new BinaryWriter(File.Create("nwd_triangles.bin")))
        using(var localWriter=new BinaryWriter(File.Create("nwd_local_triangles.bin")))
        using(var matrixWriter=new BinaryWriter(File.Create("nwd_fragment_matrices.bin"))) {
        foreach(var m in control.Document.Models) {
          Console.WriteLine("Model: "+m.FileName+" units: "+m.Units);
          models.Add(new {file=m.FileName,units=m.Units.ToString()});
          foreach(var i in m.RootItem.DescendantsAndSelf) {
            count++;
            var sceneProps=new List<object>();
            foreach(var cat in i.PropertyCategories)foreach(var prop in cat.Properties)
              if(cat.DisplayName=="Объект" || cat.DisplayName=="ID объекта" || cat.DisplayName=="Элемент")
                sceneProps.Add(new {category=cat.DisplayName,name=prop.DisplayName,value=prop.Value.ToString()});
            int parentIndex=-1;
            if(i.Parent!=null && !sceneIndices.TryGetValue(i.Parent,out parentIndex))throw new Exception("Parent missing from scene traversal");
            string sceneSource="";
            for(var ancestor=i;ancestor!=null;ancestor=ancestor.Parent)
              if(ancestor.DisplayName.EndsWith(".nwc",StringComparison.OrdinalIgnoreCase)){sceneSource=ancestor.DisplayName;break;}
            sceneIndices.Add(i,count-1);
            scene.Add(new {scene_index=count-1,parent_index=parentIndex,source=sceneSource,
              geometry_index=i.HasGeometry?geometry:-1,name=i.DisplayName,has_geometry=i.HasGeometry,properties=sceneProps});
            if(!i.HasGeometry)continue; geometry++;
            var ancestry=new List<object>();
            for(var p=i;p!=null;p=p.Parent) {
              List<object> properties;
              if(!propertiesCache.TryGetValue(p,out properties)) {
                properties=new List<object>();
                foreach(var cat in p.PropertyCategories) foreach(var prop in cat.Properties)
                  properties.Add(new {category=cat.DisplayName,name=prop.DisplayName,value=prop.Value.ToString()});
                propertiesCache[p]=properties;
              }
              ancestry.Add(new {name=p.DisplayName,properties=properties});
            }
            var callback=new Triangles { Writer=writer, LocalWriter=localWriter };
            long offset=writer.BaseStream.Position;
            var path=ComApiBridge.ToInwOaPath(i);
            int fragments=0;int skippedForeign=0;var fragmentRecords=new List<object>();
            foreach(Com.InwOaFragment3 frag in path.Fragments()) {
              var fragmentItem=ComApiBridge.ToModelItem(frag.path);
              if(!fragmentItem.Equals(i)){skippedForeign++;continue;}
              var matrix=(Array)frag.GetLocalToWorldMatrix().Matrix;
              callback.Matrix=new double[16]; int k=matrix.GetLowerBound(0);
              for(int n=0;n<16;n++) callback.Matrix[n]=Convert.ToDouble(matrix.GetValue(n+k));
              matrixWriter.Write(geometry-1);matrixWriter.Write(fragments);
              for(int n=0;n<16;n++)matrixWriter.Write(callback.Matrix[n]);
              long fragmentOffset=writer.BaseStream.Position;int beforeTriangles=callback.Count;
              frag.GenerateSimplePrimitives((Com.nwEVertexProperty)0,callback);
              string fragmentMaterial="";var materialEvidence=new List<object>();
              foreach(var cat in fragmentItem.PropertyCategories)foreach(var prop in cat.Properties)
                if((cat.DisplayName=="Материал" && prop.DisplayName=="Имя") || (cat.DisplayName=="Элемент" && prop.DisplayName=="Материал")) {
                  string value=prop.Value.ToDisplayString();
                  materialEvidence.Add(new {category=cat.DisplayName,name=prop.DisplayName,value=value});
                  if(fragmentMaterial.Length==0)fragmentMaterial=value;
                  else if(value.Length>0 && fragmentMaterial!=value)throw new Exception("Conflicting fragment materials");
                }
              fragmentRecords.Add(new {fragment_id=(geometry-1).ToString()+":"+fragments.ToString(),
                offset=fragmentOffset,triangles=callback.Count-beforeTriangles,
                local_to_world=callback.Matrix,transform_reverses=frag.IsTransformReverses,
                material=fragmentMaterial,material_evidence=materialEvidence,
                path_matches_node=fragmentItem.Equals(i),excludedBySet=excluded.Contains(fragmentItem)});
              fragments++;
            }
            var box=i.Geometry.BoundingBox;
            var item=new {index=geometry-1,offset=offset,triangles=callback.Count,fragments=fragments,skipped_foreign_instance_fragments=skippedForeign,fragmentRecords=fragmentRecords,excludedBySet=excluded.Contains(i),hidden=i.IsHidden,solid=i.Geometry.IsSolid,ancestry=ancestry,bounds=new[]{box.Min.X,box.Min.Y,box.Min.Z,box.Max.X,box.Max.Y,box.Max.Z}};
            if(!first)metadata.Write(",");first=false;metadata.Write(json.Serialize(item));
            if(geometry%1000==0) Console.WriteLine("Exported "+geometry);
          }
        }
        }
        metadata.Write("]");
        }
        File.WriteAllText("all_nodes.json",json.Serialize(new {nodes=scene}));
        File.WriteAllText("nwd_export_manifest.json",json.Serialize(new {complete=true,schema=2,
          exporter="NwdFragmentProbe-v2-literal-matrices",local_triangle_file="nwd_local_triangles.bin",local_triangle_sha256=FileHash("nwd_local_triangles.bin"),literal_matrix_file="nwd_fragment_matrices.bin",literal_matrix_sha256=FileHash("nwd_fragment_matrices.bin"),exporter_sha256=FileHash(Assembly.GetExecutingAssembly().Location),
          input_sha256=FileHash(file),input=file,models=models,read_seconds=readSeconds,extraction_seconds=watch.Elapsed.TotalSeconds-readSeconds,peak_memory_bytes=System.Diagnostics.Process.GetCurrentProcess().PeakWorkingSet64,geometry_nodes=geometry,scene_nodes=count}));
        Console.WriteLine("Items: "+count+" Geometry items: "+geometry);
      }
    } finally { ApplicationControl.Terminate(); }
    return 0;
  }
}
