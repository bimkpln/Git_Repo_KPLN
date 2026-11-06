using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
namespace KPLN_ModelChecker_Lib.Commands.Common.CheckRebarConcreteClass {
 internal static class Rules {
  internal static int? ClassFromCode(string code) {
   if(code==null || !Regex.IsMatch(code,@"\A20[0-9]{7}\z")) return null;
   return int.Parse(code.Substring(2,2),CultureInfo.InvariantCulture);
  }
  internal static decimal? ClassValue(string text) {
   decimal value;
   return !String.IsNullOrWhiteSpace(text) &&
    decimal.TryParse(text.Trim().Replace(',','.'),NumberStyles.AllowDecimalPoint|NumberStyles.AllowLeadingSign,
     CultureInfo.InvariantCulture,out value) && value>0 ? value : (decimal?)null;
  }
  internal static bool FastMode(IEnumerable<int> classes,bool materialReadComplete,int closedWorksets) {
   return materialReadComplete && closedWorksets==0 && classes.Distinct().Count()==1;
  }
  internal static string Compare(decimal? actual,int? expected,bool empty) {
   if(expected==null) return "host_class_uncertain";
   if(empty) return "rebar_class_empty";
   if(actual==null) return "rebar_class_invalid";
   return actual.Value==expected.Value ? "match" : "class_mismatch";
  }
 }
}