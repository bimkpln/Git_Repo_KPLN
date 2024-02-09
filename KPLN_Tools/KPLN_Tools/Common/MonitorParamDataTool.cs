using Autodesk.Revit.DB;
using System;

namespace KPLN_Tools.Common
{
    internal class MonitorParamDataTool
    {
        internal static double? GetDoubleValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.Double:
                    return p.AsDouble();
                case StorageType.Integer:
                    return p.AsInteger();
                case StorageType.String:
                    return double.Parse(p.AsString(), System.Globalization.NumberStyles.Float);
                default:
                    return null;
            }
        }

        internal static int? GetIntegerValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.Double:
                    return (int)Math.Round(p.AsDouble());
                case StorageType.Integer:
                    return p.AsInteger();
                case StorageType.String:
                    return int.Parse(p.AsString(), System.Globalization.NumberStyles.Integer);
                default:
                    return null;
            }
        }

        internal static string GetStringValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
                    case StorageType.ElementId:
                        return p.AsValueString();
                    case StorageType.Double:
                        return p.AsValueString();
                    case StorageType.Integer:
                        return p.AsValueString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                switch (p.StorageType)
                {
                    case StorageType.ElementId:
                        return p.AsValueString();
                    case StorageType.Double:
                        return p.AsDouble().ToString();
                    case StorageType.Integer:
                        return p.AsInteger().ToString();
                    case StorageType.String:
                        return p.AsString();
                    default:
                        return null;
                }
            }
        }
    }
}
