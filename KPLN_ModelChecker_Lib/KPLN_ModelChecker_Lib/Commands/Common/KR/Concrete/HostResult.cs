using System.Collections.Generic;
using Autodesk.Revit.DB;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Concrete {
 internal sealed class HostResult {
  internal string Status;
  internal string Mode="geometry";
  internal ElementId Id;
  internal int? Class;
  internal readonly List<string> Notes=new List<string>();
  internal readonly List<Dictionary<string,object>> Candidates=new List<Dictionary<string,object>>();
 }
}
