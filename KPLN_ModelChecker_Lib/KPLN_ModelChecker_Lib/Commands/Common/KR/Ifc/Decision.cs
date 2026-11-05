using System;
namespace KPLN_ModelChecker_Lib.Commands.Common.KR.Ifc {
 public static class Decision {
  // Input order resolves an exact tie. Caller orders hosts by numeric ElementId.
  public static int LargestIntersectionIndex(double[] volumes) {
   int best=-1;
   for(int i=0;i<volumes.Length;i++) {
    if(double.IsNaN(volumes[i]) || double.IsInfinity(volumes[i]) || volumes[i]<=0)
     throw new ArgumentException("Expected positive finite intersection volumes.");
    if(best<0 || volumes[i]>volumes[best]) best=i;
   }
   return best;
  }
  public static string Evaluate(int hosts, bool uncertain, bool parameterExists, string actual, string expected) {
   if (uncertain) return "geometry_uncertain";
   if (hosts == 0) return "host_not_found";
   if (!parameterExists) return "parameter_missing";
   if (String.IsNullOrWhiteSpace(expected)) return "host_mark_empty";
   if (String.IsNullOrWhiteSpace(actual)) return "rebar_mark_empty";
   return String.Equals(actual, expected, StringComparison.Ordinal) ? "match" : "mark_mismatch";
  }
 }
}