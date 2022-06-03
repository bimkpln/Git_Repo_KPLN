using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_ModelChecker_User.Tools
{
    public static class Extentions
    {
        private readonly static string[] stopList = new string[] { "199_", "501_" };
        public static bool ElementPassesConditions(this Element element)
        {
            try
            {
                if (element.GetType() == typeof(FamilyInstance))
                {
                    foreach (string i in stopList)
                    {
                        if ((element as FamilyInstance).Symbol.FamilyName.StartsWith(i)) return false;
                        if ((element as FamilyInstance).Symbol.Name.StartsWith(i)) return false;
                    }
                }
                else
                {
                    foreach (string i in stopList)
                    {
                        if (element.Name.StartsWith(i)) return false;
                    }
                }
            }
            catch (Exception) { }
            return true;
        }
    }
}
