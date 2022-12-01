using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace KPLN_Classificator
{
    [Serializable]
    public class Classificator
    {
        public List<string> paramsValues;
        public BuiltInCategory BuiltInName;
        public string FamilyName;
        public string TypeName;

        public Classificator()
        {

        }
    }
}
