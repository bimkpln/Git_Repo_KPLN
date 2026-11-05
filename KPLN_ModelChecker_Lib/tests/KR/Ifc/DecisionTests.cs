using System;
using KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc;
class DecisionTests {
 static int count;
 static void Expect(string expected,int hosts,bool uncertain,bool parameter,string actual,string mark) {
  var result=Decision.Evaluate(hosts,uncertain,parameter,actual,mark);
  if(result!=expected) throw new Exception(expected+" != "+result);
  count++;
 }
 static void Main() {
  Expect("match",1,false,true,"Фп13.1-1","Фп13.1-1");
  Expect("mark_mismatch",1,false,true,"Фп13.1-2","Фп13.1-1");
  Expect("mark_mismatch",1,false,true,"Фп13.1-1 ","Фп13.1-1");
  Expect("mark_mismatch",1,false,true,"фп13.1-1","Фп13.1-1");
  Expect("rebar_mark_empty",1,false,true," ","Фп13.1-1");
  Expect("host_mark_empty",1,false,true,"Фп13.1-1",null);
  Expect("host_mark_empty",1,false,true,null,"");
  Expect("parameter_missing",1,false,false,null,"Фп13.1-1");
  Expect("host_not_found",0,false,true,"Фп13.1-1",null);
  Expect("match",2,false,true,"Фп13.1-1","Фп13.1-1");
  Expect("mark_mismatch",3,false,true,"Фп13.1-2","Фп13.1-1");
  Expect("geometry_uncertain",2,true,true,"Фп13.1-1","Фп13.1-1");
  var volumes=new[]{0.1,0.8,0.3};
  if(Decision.LargestIntersectionIndex(volumes)!=1) throw new Exception("Largest volume was not selected");
  if(Decision.LargestIntersectionIndex(new[]{0.8,0.1})!=0) throw new Exception("First largest lost");
  if(Decision.LargestIntersectionIndex(new[]{0.1,0.8})!=1) throw new Exception("Last largest lost");
  if(Decision.LargestIntersectionIndex(new[]{0.8,0.8})!=0) throw new Exception("Tie must be stable");
  if(Decision.LargestIntersectionIndex(new double[0])!=-1) throw new Exception("Empty host set");
  if(Decision.LargestIntersectionIndex(new[]{0.2})!=0) throw new Exception("Single host");
  var marks=new[]{"Wrong","Фп13.1-1","Wrong"};
  Expect("match",3,false,true,"Фп13.1-1",marks[Decision.LargestIntersectionIndex(volumes)]);
  foreach(var invalid in new[]{double.NaN,double.PositiveInfinity,0.0,-1.0}) {
   try { Decision.LargestIntersectionIndex(new[]{invalid}); throw new Exception("Invalid volume accepted"); }
   catch(ArgumentException) {}
  }
  count+=10;
  Expect("geometry_uncertain",1,true,true,"Фп13.1-1","Фп13.1-1");
  Expect("geometry_uncertain",0,true,false,null,null);
  Console.WriteLine("PASS: "+count+" mark/host decisions, including uncertainty precedence.");
 }
}