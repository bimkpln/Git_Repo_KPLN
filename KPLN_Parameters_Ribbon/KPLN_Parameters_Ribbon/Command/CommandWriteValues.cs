using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Parameters_Ribbon.Common;
using KPLN_Parameters_Ribbon.Forms;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon.Command
{
    public class CommandWriteValues : IExecutableCommand
    {
        private List<ParameterRuleElement> Rules { get; set; }
        public CommandWriteValues(ObservableCollection<ParameterRuleElement> rules)
        {
            Rules = rules.ToList();
        }
        private Parameter GetParameterByElement(Element element, ListBoxElement rule)
        {
            foreach (Parameter p in element.Parameters)
            {
                if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                {
                    return p;
                }
            }
            try
            {
                foreach (Parameter p in element.Document.GetElement(element.GetTypeId()).Parameters)
                {
                    if (p.Definition.Name == (rule.Data as Parameter).Definition.Name)
                    {
                        return p;
                    }
                }
            }
            catch (Exception)
            { }
            return null;
        }
        public Result Execute(UIApplication app)
        {
            try
            {
                Document doc = app.ActiveUIDocument.Document;
                int max = 0;
                foreach (ParameterRuleElement rule in Rules)
                {
                    if (rule.SelectedCategory.Data == null || rule.SelectedSourceParameter == null || rule.SelectedTargetParameter == null) { continue; }
                    max += new FilteredElementCollector(doc).OfCategoryId((rule.SelectedCategory.Data as Category).Id).ToElements().Count;
                }
                string format = "{0} из " + max.ToString() + " элементов обработано";
                using (Progress_Single pb = new Progress_Single("Копирование параметров", format, max))
                {
                    using (Transaction t = new Transaction(doc, "Копирование параметров"))
                    {
                        t.Start();
                        try
                        {
                            foreach (ParameterRuleElement rule in Rules)
                            {
                                try
                                {
                                    if (rule.SelectedCategory.Data == null || rule.SelectedSourceParameter == null || rule.SelectedTargetParameter == null) { continue; }
                                    foreach (Element element in new FilteredElementCollector(doc).OfCategoryId((rule.SelectedCategory.Data as Category).Id).ToElements())
                                    {
                                        pb.Increment();
                                        try
                                        {
                                            Parameter sourceParameter = GetParameterByElement(element, rule.SelectedSourceParameter);
                                            Parameter targetParameter = GetParameterByElement(element, rule.SelectedTargetParameter);
                                            if (sourceParameter != null && targetParameter != null)
                                            {
                                                switch (targetParameter.StorageType)
                                                {
                                                    case StorageType.Double:
                                                        double? dv = GetDoubleValue(sourceParameter);
                                                        if (dv != null)
                                                        {
                                                            targetParameter.Set((double)dv);
                                                        }
                                                        break;
                                                    case StorageType.Integer:
                                                        int? iv = GetIntegerValue(sourceParameter);
                                                        if (iv != null)
                                                        {
                                                            targetParameter.Set((int)iv);
                                                        }
                                                        break;
                                                    case StorageType.String:
                                                        string sv = GetStringValue(sourceParameter);
                                                        if (sv != null && sv != " " && sv != string.Empty)
                                                        {
                                                            targetParameter.Set(sv);
                                                        }
                                                        break;
                                                }
                                            }
                                        }
                                        catch (Exception) { continue; }
                                    }
                                }
                                catch (Exception) { continue; }
                            }
                            t.Commit();
                        }
                        catch (Exception) { t.RollBack(); }
                    }
                }
                return Result.Succeeded;
            }
            catch (Exception)
            {
                return Result.Failed;
            }
        }
        public double? GetDoubleValue(Parameter p)
        {
            try
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
            catch (Exception) { return null; }
        }
        public int? GetIntegerValue(Parameter p)
        {
            try
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
            catch (Exception)
            {
                return null;
            }
        }
        public string GetStringValue(Parameter p)
        {
            try
            {
                switch (p.StorageType)
                {
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
                try
                {
                    switch (p.StorageType)
                    {
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
                catch (Exception) { return null; }
            }
        }
    }
}
