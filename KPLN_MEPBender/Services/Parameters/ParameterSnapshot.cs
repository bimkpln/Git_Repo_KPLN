using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace KPLN_MEPBender.Services.Parameters
{
    public sealed class ParameterSnapshot
    {
        public ParameterSnapshot()
        {
            Values = new Dictionary<string, ParameterSnapshotValue>();
        }

        public Dictionary<string, ParameterSnapshotValue> Values { get; }
    }

    public sealed class ParameterSnapshotValue
    {
        public string Key { get; set; }

        public string Name { get; set; }

        public StorageType StorageType { get; set; }

        public int? ParameterIdInteger { get; set; }

        public BuiltInParameter? BuiltInParameter { get; set; }

        public string StringValue { get; set; }

        public double DoubleValue { get; set; }

        public int IntegerValue { get; set; }

        public ElementId ElementIdValue { get; set; }
    }
}
