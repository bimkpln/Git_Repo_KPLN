using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows;
//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.DB.Mechanical;
//using Autodesk.Revit.DB.Plumbing;
//using KPLN_Tools.ExternalCommands.Specification.Specification;
//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Utils;

namespace KPLN_Tools.ExternalCommands
{
    namespace Specification
    {
        internal class Command_SET_ElementsNotModeled : IExternalCommand
        {
            public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
            {
                MessageBox.Show("Используй плагин от СМЛТ. Этот пока не активен");
                return Result.Cancelled;
            }
        }
    }
}


//        public class ParametersSM
//        {
//            public CategorySet Get(Document doc, string sharedParameterName)
//            {
//                DefinitionBindingMapIterator bindingMapIterator = ((DefinitionBindingMap)doc.ParameterBindings).ForwardIterator();
//                while (bindingMapIterator.MoveNext())
//                {
//                    Definition key = bindingMapIterator.Key;
//                    if (sharedParameterName.Equals(key.Name, StringComparison.CurrentCultureIgnoreCase))
//                        return ((ElementBinding)bindingMapIterator.Current).Categories;
//                }
//                return (CategorySet)null;
//            }

//            public bool CheckInCategoties(Document doc, IList<string> parameterNames, string checkCat)
//            {
//                List<string> values = new List<string>();
//                foreach (string parameterName in (IEnumerable<string>)parameterNames)
//                {
//                    bool flag = false;
//                    foreach (Category category in this.Get(doc, parameterName))
//                    {
//                        if (category.Name == checkCat)
//                            flag = true;
//                    }
//                    if (!flag)
//                        values.Add("Параметр " + parameterName + " не назнчаен категории " + checkCat + ".");
//                }
//                if (values.Count <= 0)
//                    return true;
//                TaskDialog.Show("Ошибка!", string.Join("\n", (IEnumerable<string>)values) + "\nПроверьте параметры проекта.");
//                return false;
//            }

//            public ElementId GetParametrId(Document doc, string name)
//            {
//                return new FilteredElementCollector(doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == name)).FirstOrDefault<Element>().Id;
//            }
//        }

//        public class ParameterFilters
//        {
//            public ElementParameterFilter GetFilterStringEquals(ElementId parameterId, string value)
//            {
//                return new ElementParameterFilter((FilterRule)new FilterStringRule((FilterableValueProvider)new ParameterValueProvider(parameterId), (FilterStringRuleEvaluator)new FilterStringEquals(), value, true));
//            }

//            public ElementParameterFilter GetFilterStringNotEquals(ElementId parameterId, string value)
//            {
//                return new ElementParameterFilter((FilterRule)new FilterInverseRule((FilterRule)new FilterStringRule((FilterableValueProvider)new ParameterValueProvider(parameterId), (FilterStringRuleEvaluator)new FilterStringEquals(), value, true)));
//            }

//            public ElementParameterFilter GetFilterDoubleNotEquals(ElementId parameterId, double value)
//            {
//                return new ElementParameterFilter((FilterRule)new FilterDoubleRule((FilterableValueProvider)new ParameterValueProvider(parameterId), (FilterNumericRuleEvaluator)new FilterNumericGreater(), value, 1E-06));
//            }

//            public ElementParameterFilter GetFilterDoubleEquals(ElementId parameterId, double value)
//            {
//                return new ElementParameterFilter((FilterRule)new FilterInverseRule((FilterRule)new FilterDoubleRule((FilterableValueProvider)new ParameterValueProvider(parameterId), (FilterNumericRuleEvaluator)new FilterNumericGreater(), value, 1E-06)));
//            }
//        }

//        public class EstimateParameters
//        {
//            public Guid GuidParamFolder { get; set; }

//            public Guid GuidParamSection { get; set; }

//            public Guid GuidParamLevel { get; set; }

//            public Guid GuidParamDiscipline { get; set; }

//            public Guid GuidParamChapter { get; set; }

//            public Guid GuidParamSmeta { get; set; }

//            public Guid GuidParamCount { get; set; }

//            public Guid GuidParamDnar { get; set; }

//            public Guid GuidParamSizeLanthFitt { get; set; }

//            public Guid GuidParamSizeAreaFitt { get; set; }

//            public Guid GuidParamDy { get; set; }

//            public ElementId idParamFolder { get; set; }

//            public ElementId idParamLevel { get; set; }

//            public ElementId idParamSection { get; set; }

//            public void Get(Document _doc)
//            {
//                this.idParamFolder = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Папка Работ")).FirstOrDefault<Element>().Id;
//                this.GuidParamFolder = (_doc.GetElement(this.idParamFolder) as SharedParameterElement).GuidValue;
//                this.idParamSection = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Секция")).FirstOrDefault<Element>().Id;
//                this.GuidParamSection = (_doc.GetElement(this.idParamSection) as SharedParameterElement).GuidValue;
//                this.idParamLevel = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Этаж")).FirstOrDefault<Element>().Id;
//                this.GuidParamLevel = (_doc.GetElement(this.idParamLevel) as SharedParameterElement).GuidValue;
//                ElementId id1 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Дисциплина")).FirstOrDefault<Element>().Id;
//                this.GuidParamDiscipline = (_doc.GetElement(id1) as SharedParameterElement).GuidValue;
//                ElementId id2 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Раздел")).FirstOrDefault<Element>().Id;
//                this.GuidParamChapter = (_doc.GetElement(id2) as SharedParameterElement).GuidValue;
//                ElementId id3 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "СМ_Смета")).FirstOrDefault<Element>().Id;
//                this.GuidParamSmeta = (_doc.GetElement(id3) as SharedParameterElement).GuidValue;
//                ElementId id4 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "ASML_Количество")).FirstOrDefault<Element>().Id;
//                this.GuidParamCount = (_doc.GetElement(id4) as SharedParameterElement).GuidValue;
//                try
//                {
//                    ElementId id5 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "ASML_Диаметр наружный")).FirstOrDefault<Element>().Id;
//                    this.GuidParamDnar = (_doc.GetElement(id5) as SharedParameterElement).GuidValue;
//                }
//                catch
//                {
//                }
//                try
//                {
//                    ElementId id6 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "ASML_Размер_Длина фитинга")).FirstOrDefault<Element>().Id;
//                    this.GuidParamSizeLanthFitt = (_doc.GetElement(id6) as SharedParameterElement).GuidValue;
//                }
//                catch
//                {
//                }
//                try
//                {
//                    ElementId id7 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "ASML_Размер_Площадь фитинга")).FirstOrDefault<Element>().Id;
//                    this.GuidParamSizeAreaFitt = (_doc.GetElement(id7) as SharedParameterElement).GuidValue;
//                }
//                catch
//                {
//                }
//                try
//                {
//                    ElementId id8 = new FilteredElementCollector(_doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == "ASML_Диаметр условный")).FirstOrDefault<Element>().Id;
//                    this.GuidParamDy = (_doc.GetElement(id8) as SharedParameterElement).GuidValue;
//                }
//                catch
//                {
//                }
//            }

//            public bool CheckParameters(EstimateParameters estParam, IList<Element> pipesIsFitting)
//            {
//                IList<Element> elementList = (IList<Element>)new List<Element>();
//                IList<Element> list = (IList<Element>)pipesIsFitting.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt) == null || a.get_Parameter(estParam.GuidParamDy) == null)).ToList<Element>();
//                if (list.Count <= 0)
//                    return false;
//                TaskDialog.Show("Ошибка!", "Параметры ASML_Размер_Длина фитинга или ASML_Диаметр условный не добавлены в семейство или располагаются в типе, имена семейств и типоразмеров:\n\n" + string.Join("\n", list.Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() + "-->" + a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()))) + "\n\nОбновите семейства и повторите запуск. При повторной ошибке обратитесь в отдел BIM для добавления параметров в семейства, переноса параметров в экземляр. Расчет приостановлен.");
//                return true;
//            }

//            public bool CheckParametersReadOnly(
//              Document Doc,
//              EstimateParameters estParam,
//              FamilySymbol symbol)
//            {
//                List<string> stringList = new List<string>();
//                this.Check(Doc, symbol, estParam.GuidParamCount, stringList);
//                this.Check(Doc, symbol, estParam.GuidParamFolder, stringList);
//                this.Check(Doc, symbol, estParam.GuidParamSection, stringList);
//                this.Check(Doc, symbol, estParam.GuidParamLevel, stringList);
//                this.Check(Doc, symbol, estParam.GuidParamDiscipline, stringList);
//                this.Check(Doc, symbol, estParam.GuidParamChapter, stringList);
//                if (stringList.Count<string>() <= 0)
//                    return true;
//                TaskDialog.Show("Ошибка!", "В семейство ASML_ОВ_ВК_Элементы2D \n" + string.Join("\n", (IEnumerable<string>)stringList) + ". \nУдалите параметр(ы) из типа.");
//                return false;
//            }

//            public void Check(Document Doc, FamilySymbol symbol, Guid guid, List<string> checks)
//            {
//                if (((Element)symbol).LookupParameter(this.GetParameterNameByGuid(Doc, guid)) == null)
//                    return;
//                checks.Add("добавлен в типы или заблокирован параметр " + this.GetParameterNameByGuid(Doc, guid));
//            }

//            public string GetParameterNameByGuid(Document doc, Guid paramGuid)
//            {
//                SharedParameterElement parameterElement = SharedParameterElement.Lookup(doc, paramGuid);
//                return parameterElement != null ? ((Definition)(doc.GetElement(((Element)parameterElement).Id) as ParameterElement).GetDefinition()).Name : (string)null;
//            }
//        }

//        namespace Specification
//        {
//            public class Calcs
//            {
//                internal double GetPipesMetallSumIso(IList<Element> ListP)
//                {
//                    return (ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 17.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 2.5 * 0.34))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 17.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 23.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 3.0 * 0.36))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 23.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 27.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 3.5 * 0.37))
//                            .Sum()
//                            +
//                            ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 27.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 35.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 4.0 * 0.4))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 35.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 45.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 4.5 * 0.51))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 45.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 55.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 5.0 * 0.72))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 55.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 70.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 6.0 * 1.17))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 70.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 90.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 6.0 * 1.34))
//                            .Sum()
//                            +
//                            ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 90.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 115.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 6.0 * 2.23))
//                            .Sum()
//                            +
//                            ListP.
//                            Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 115.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 7.0 * 3.7))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 8.0 * 4.92))
//                            .Sum()) / 1000.0 / 1000.0;
//                }

//                internal double GetFittMetallSumIso(
//                  Document doc,
//                  EstimateParameters estParam,
//                  IList<Element> ListP)
//                {
//                    return (ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 17.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 2.5 * 0.34))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 17.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 23.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 3.0 * 0.36))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 23.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 27.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 3.5 * 0.37))
//                            .Sum()
//                            +
//                            ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 27.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 35.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 4.0 * 0.4))
//                            .Sum()
//                            +
//                            ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 35.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 45.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 4.5 * 0.51))
//                            .Sum()
//                            +
//                            ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 45.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 55.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 5.0 * 0.72))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 55.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 70.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 6.0 * 1.17))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 70.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 90.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 6.0 * 1.34))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 90.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 115.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 6.0 * 2.23))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 115.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 7.0 * 3.7))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 8.0 * 4.92))
//                            .Sum()) / 1000.0 / 1000.0;
//                }

//                internal double GetCleiPipe(IList<Element> ListP)
//                {
//                    return ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 1.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 7.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 270.0 * 1.25)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 7.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 10.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 180.0 * 1.25)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 10.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 14.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 125.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 14.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 21.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 80.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 21.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 26.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 55.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 26.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 33.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 45.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 33.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 41.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 35.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 41.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1000.0 / 4.5 * 1.5)).Sum();
//                }

//                internal double GetCleiFitt(EstimateParameters estParam, IList<Element> ListP)
//                {
//                    return ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 1.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 7.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 270.0 * 1.25)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 7.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 10.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 180.0 * 1.25)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 10.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 14.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 125.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 14.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 21.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 80.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 21.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 26.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 55.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 26.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 33.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 45.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 33.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 <= 41.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 35.0 * 1.2)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() * 304.8 > 41.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1000.0 / 4.5 * 1.5)).Sum();
//                }

//                internal double GetMetallSumIsoDuct(IList<Element> ListP)
//                {
//                    return (ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 <= 160.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 0.33)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 > 160.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 <= 315.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 0.75)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 > 315.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 <= 500.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 1.8)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 > 500.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 <= 700.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 4.0)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 > 700.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 <= 900.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 6.5)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_EQ_DIAMETER_PARAM).AsDouble() * 304.8 > 900.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 * 8.8)).Sum()) / 1000.0 / 1000.0;
//                }

//                internal double GetPipesMetallSum(IList<Element> ListP)
//                {
//                    return (ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 17.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 1.5 * 1.09))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 17.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 23.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 2.0 * 1.1))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 23.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 27.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 2.0 * 1.25))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 27.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 35.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 2.5 * 1.75))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 35.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 45.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 3.0 * 2.18))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 45.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 55.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 3.0 * 2.43))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 55.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 70.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 4.0 * 3.4))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 70.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 90.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 4.0 * 4.3))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 90.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 115.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 4.5 * 6.0))
//                            .Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 115.0))
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 < 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 5.0 * 7.4))
//                            .Sum()
//                            +
//                            ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble() * 304.8 >= 135.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() * 304.8 / 6.0 * 10.6))
//                            .Sum()) / 1000.0 / 1000.0;
//                }

//                internal double GetFittMetallSum(
//                  Document doc,
//                  EstimateParameters estParam,
//                  IList<Element> ListP)
//                {
//                    return (ListP
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 17.0))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 1.5 * 1.09))
//                            .Sum()
//                            + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 17.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 23.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 2.0 * 1.1)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 23.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 27.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 2.0 * 1.25)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 27.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 35.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 2.5 * 1.75)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 35.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 45.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 3.0 * 2.18)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 45.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 55.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 3.0 * 2.43)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 55.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 70.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 4.0 * 3.4)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 70.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 90.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 4.0 * 4.3)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 90.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 115.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 4.5 * 6.0)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 115.0)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 < 135.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 5.0 * 7.4)).Sum() + ListP.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamDy).AsDouble() * 304.8 >= 135.0)).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 / 6.0 * 10.6)).Sum()) / 1000.0 / 1000.0;
//                }

//                internal double GetDuctGlueSum(IList<Element> elements)
//                {
//                    double koef = 10.7639104167097;
//                    return elements
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI30")))
//                            .Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 0.6))
//                            .Sum()
//                            + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI60"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 0.92)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI90"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 0.6)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI120"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 1.52)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI150"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 2.05)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI180"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 3.05)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI240"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.RBS_CURVE_SURFACE_AREA).AsDouble() / koef * 3.05)).Sum();
//                }

//                internal double GetDuctFittingGlueSum(
//                  Document doc,
//                  EstimateParameters estParam,
//                  IList<Element> elements)
//                {
//                    double koef = 10.7639104167097;
//                    return elements
//                            .Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI30")))
//                            .Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 0.6))
//                            .Sum()
//                            + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI60"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 0.92)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI90"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 0.6)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI120"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 1.52)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI150"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 2.05)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI180"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 3.05)).Sum() + elements.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString().Contains("EI240"))).Select<Element, double>((Func<Element, double>)(a => doc.GetElement((a as InsulationLiningBase).HostElementId).get_Parameter(estParam.GuidParamSizeAreaFitt).AsDouble() / koef * 3.05)).Sum();
//                }
//            }
//        }

//        public class ParametersDto
//        {
//            public Document doc { get; set; }

//            public string message { get; set; }

//            public Guid guidParamCount { get; set; }

//            public Guid guidParamName { get; set; }

//            public Guid guidParamThicknessPipe { get; set; }

//            public Guid guidParamNamePipeAndDuct { get; set; }

//            public Guid guidParamType { get; set; }

//            public Guid guidParamMark { get; set; }

//            public Guid guidParamSizeArea { get; set; }

//            public Guid guidParamSizeUnit { get; set; }

//            public bool exit { get; set; }

//            public List<string> _erroPipeSegmet { get; set; }

//            public List<Element> _pipeWithNotOneSegmet { get; set; }
//        }
//        public class Parameters
//        {
//            public (List<string>, bool) CheckNull(in List<string> strings, string pramName, string cat)
//            {
//                bool flag = false;
//                if (strings.Where<string>((Func<string, bool>)(a => a == null)).Count<string>() > 0)
//                    flag = true;
//                else if (strings.Where<string>((Func<string, bool>)(a => a == "" || a == " ")).Count<string>() > 0)
//                    flag = true;
//                List<string> stringList = new List<string>();
//                if (!flag)
//                    return (strings, flag);
//                try
//                {
//                    if (MessageBox.Show("В категории " + cat + " не заполнен параметр " + pramName + ". Продолжить выполнение программы?", "Предупреждение!", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
//                        return (stringList, true);
//                    stringList = strings.Where<string>((Func<string, bool>)(a => a != null)).Where<string>((Func<string, bool>)(a => a != "" || a != " ")).ToList<string>();
//                    return (stringList, false);
//                }
//                catch
//                {
//                    return (stringList, true);
//                }
//            }

//            public ParametersDto Get(ParametersDto parametersDto)
//            {
//                parametersDto.message = "";
//                parametersDto.exit = false;
//                parametersDto.guidParamCount = this.CheckAvailability(parametersDto, "ASML_Количество");
//                parametersDto.guidParamNamePipeAndDuct = this.CheckAvailability(parametersDto, "ASML_Наименование трубы и воздуховода");
//                parametersDto.guidParamType = this.CheckAvailability(parametersDto, "ASML_Тип");
//                parametersDto.guidParamMark = this.CheckAvailability(parametersDto, "ASML_Марка");
//                parametersDto.guidParamName = this.CheckAvailability(parametersDto, "ASML_Наименование");
//                parametersDto.guidParamSizeArea = this.CheckAvailability(parametersDto, "ASML_Размер_Площадь фитинга");
//                parametersDto.guidParamSizeUnit = this.CheckAvailability(parametersDto, "ASML_Единица измерения");
//                return parametersDto;
//            }

//            public Guid CheckAvailability(ParametersDto parametersDto, string name)
//            {
//                try
//                {
//                    IEnumerable<Element> source = new FilteredElementCollector(parametersDto.doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == name));
//                    if (source.Count<Element>() > 1)
//                    {
//                        parametersDto.message = parametersDto.message + "Удалите дублирование параметра \n" + name + "\nId параметров\n" + string.Join<ElementId>("\n", source.Select<Element, ElementId>((Func<Element, ElementId>)(a => a.Id))) + "\n\n";
//                        return new Guid();
//                    }
//                    ElementId id = new FilteredElementCollector(parametersDto.doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == name)).FirstOrDefault<Element>().Id;
//                    return (parametersDto.doc.GetElement(id) as SharedParameterElement).GuidValue;
//                }
//                catch
//                {
//                    parametersDto.message = parametersDto.message + "Отсутствует параметр " + name + "\n";
//                    return new Guid();
//                }
//            }
//        }


//        [Transaction(TransactionMode.Manual)]
//        [Regeneration(RegenerationOption.Manual)]
//        internal class Command_SET_ElementsNotModeled : IExternalCommand
//        {
//            public static string commandVersion = "1.9.0.0";
//            public Document Doc;
//            public FamilyInstance ElementInstance = (FamilyInstance)null;
//            public double Quantity;
//            public string FolderJob;
//            public string Section;
//            public string Levels;
//            public string Discipline;
//            public string Chapter;
//            public Calcs Calc;
//            public double Ij = 1.0;
//            public FamilySymbol SymbolMetPipes = (FamilySymbol)null;
//            public FamilySymbol SymbolMetDuct = (FamilySymbol)null;
//            public FamilySymbol SymbolClei = (FamilySymbol)null;
//            public Level LevMax = (Level)null;
//            public bool CheckParams = false;
//            public ParameterFilters _parameterFilter = new ParameterFilters();
//            public List<string> _insulationPipeNotClei = new List<string>();
//            public Parameters parameters = new Parameters();

//            public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
//            {
//                this.Calc = new Calcs();
//                this.Doc = commandData.Application.ActiveUIDocument.Document;
//                EstimateParameters estParam = new EstimateParameters();
//                estParam.Get(this.Doc);
//                this._insulationPipeNotClei = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeInsulationType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != "Не рассчитывать клей")).Select<Element, string>((Func<Element, string>)(a => a.Name)).ToList<string>();
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Металл для крепления_ОВ_ВК")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Металл для крепления_ОВ_ВК")).FirstOrDefault<Element>().LookupParameter("ASML_Единица измерения").AsString() != "т")
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                if ((((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Металл для крепления_ОВ_ВК")).FirstOrDefault<Element>() as ElementType).FamilyName != "ASML_ОВ_ВК_Элементы2D")
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }

//                if (UnitsProject.Check(this.Doc))
//                    return (Result)0;
//                Stopwatch stopwatch = new Stopwatch();
//                stopwatch.Start();
//                if (MessageBox.Show("Перед запуском проверьте и заполните:\nпараметры СМ_Папка Работ, СМ_Этаж, СМ_Секция\n\nПродолжить?", "Предупреждение!", MessageBoxButton.YesNo) == MessageBoxResult.No)
//                    return (Result)0;
//                IEnumerable<Element> source1 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(Pipe)).WhereElementIsNotElementType()).Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe)).Name.ToString().Contains("ASML_ОВ_Медь Сварка")));
//                bool flag = false;
//                if (source1.Count<Element>() != 0 && MessageBox.Show("В проекте обнаружены трубы для кондициоонировани ASML_ОВ_Медь Сварка, расчитать эл.кабель для данного типа труб?", "Предупреждение!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
//                    flag = true;
//                IList<BuiltInCategory> builtInCategoryList = (IList<BuiltInCategory>)new List<BuiltInCategory>()
//                {
//                BuiltInCategory.OST_MechanicalEquipment,
//                BuiltInCategory.OST_Sprinklers,
//                BuiltInCategory.OST_PlumbingFixtures,
//                BuiltInCategory.OST_DuctTerminal,
//                BuiltInCategory.OST_DuctAccessory,
//                BuiltInCategory.OST_PipeAccessory,
//                BuiltInCategory.OST_PipeCurves,
//                BuiltInCategory.OST_DuctCurves,
//                BuiltInCategory.OST_FlexPipeCurves,
//                BuiltInCategory.OST_FlexDuctCurves,
//                BuiltInCategory.OST_PipeFitting,
//                BuiltInCategory.OST_DuctFitting,
//                BuiltInCategory.OST_PipeInsulations,
//                BuiltInCategory.OST_DuctInsulations,
//                BuiltInCategory.OST_StructuralFraming,
//                BuiltInCategory.OST_GenericModel,
//                BuiltInCategory.OST_TelephoneDevices,
//                BuiltInCategory.OST_ElectricalEquipment
//                };
//                IList<string> parameterNames = (IList<string>)new List<string>()
//                {
//                "СМ_Папка Работ",
//                "СМ_Этаж",
//                "СМ_Секция",
//                "СМ_Дисциплина",
//                "СМ_Раздел",
//                "СМ_Смета"
//                };
//                IList<string> stringList1 = (IList<string>)new List<string>();
//                IList<Guid> guidList = (IList<Guid>)new List<Guid>();
//                foreach (string str in (IEnumerable<string>)parameterNames)
//                {
//                    string item = str;
//                    IEnumerable<Element> source2 = new FilteredElementCollector(this.Doc).OfClass(typeof(ParameterElement)).ToElements().Where<Element>((Func<Element, bool>)(a => a.Name == item));
//                    if (source2.Count<Element>() == 0)
//                        stringList1.Add("Добавьте параметр " + item + " в проект");
//                    else if (source2.Count<Element>() > 1)
//                        stringList1.Add("Удалите дублирование параметра " + item);
//                    else if (source2.Count<Element>() == 1)
//                        guidList.Add((this.Doc.GetElement(source2.FirstOrDefault<Element>().Id) as SharedParameterElement).GuidValue);
//                }
//                if (!new ParametersSM().CheckInCategoties(this.Doc, parameterNames, "Телефонные устройства"))
//                    return (Result)0;
//                ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).WhereElementIsNotElementType()).Where<Element>((Func<Element, bool>)(a => a.Category != null));
//                foreach (Guid guid1 in (IEnumerable<Guid>)guidList)
//                {
//                    Guid guid = guid1;
//                    foreach (BuiltInCategory builtInCategory in (IEnumerable<BuiltInCategory>)builtInCategoryList)
//                    {
//                        Element element = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).WhereElementIsNotElementType().OfCategory(builtInCategory)).FirstOrDefault<Element>();
//                        if (element != null && element.get_Parameter(guid) == null)
//                        {
//                            string name = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(SharedParameterElement))).Where<Element>((Func<Element, bool>)(a => (a as SharedParameterElement).GuidValue.ToString() == guid.ToString())).FirstOrDefault<Element>().Name;
//                            stringList1.Add("Отсутствует параметр " + name + " в категории " + element.Category.Name.ToString());
//                        }
//                    }
//                }
//                if (stringList1.Count<string>() > 0)
//                {
//                    TaskDialog.Show("Ошибка!", string.Join("\n", (IEnumerable<string>)stringList1));
//                    return (Result)0;
//                }
//                ElementId id1 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).First<Element>().LookupParameter("Имя типа").Id;
//                ElementId parameterId1 = new ElementId(BuiltInParameter.SYMBOL_NAME_PARAM);
//                ElementId parameterId2 = new ElementId(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS);
//                ElementId id2 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).First<Element>().get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).Id;
//                IList<string> list1 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list2 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list3 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() == "Грунтовка_Краска_ВК")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list4 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() == "Грунтовка_Краска_ОВ")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list5 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() == "Крепления_НЕТ")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list6 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() == "Цинол_ОВ_ВК_ИТП")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                IList<string> list7 = (IList<string>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeType))).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString().Length > 3)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString() == "Краска_ИТП")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString())).ToList<string>();
//                if (list2.Count<string>() - list3.Count<string>() - list4.Count<string>() - list5.Count<string>() - list6.Count<string>() - list7.Count<string>() != 0)
//                {
//                    TaskDialog.Show("Предупреждение!", "В параметре (Комментарии к типоразмеру) в ТИПАХ (неиспользуемые включительно) труб неверно заполнены или не заполнены значения, проверьте пробелы и регистры в значениях Грунтовка_Краска_ВК, Грунтовка_Краска_ОВ, Краска_ИТП, Крепления_НЕТ, Цинол_ОВ_ВК_ИТП, см. таблицы параметров труб");
//                    return (Result)0;
//                }
//                if (!(Path.GetFileName(this.Doc.PathName).Contains("_ИТП") | Path.GetFileName(this.Doc.PathName).Contains("_ТМ")) && ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilyInstance))).Where<Element>((Func<Element, bool>)(x => x.Name == "ASML_ОВ_Гильза" | x.Name == "ASML_ВК_Гильза")).Count<Element>() == 0)
//                {
//                    switch (MessageBox.Show("В проекте отсутствуют гильзы или применены неактуальные семейства!\nНаименование типоразмеров гильз используемые плагином для расчета:\nASML_ОВ_Гильза\nASML_ВК_Гильза\n\nПродолжить расчет без гильз?", "Предупреждение!", MessageBoxButton.YesNo))
//                    {
//                        case MessageBoxResult.No:
//                            return (Result)0;
//                    }
//                }
//                FamilySymbol familySymbol1 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Цинол_ОВ_ВК_ИТП")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Цинол_ОВ_ВК_ИТП")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol2 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol3 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_З")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_З")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol4 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_К")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_К")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol5 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_Г")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ВК_АПТ_Г")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol6 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol7 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ИТП")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Краска_ИТП")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol8 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка_ВК")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol9 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol10 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка гильз_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка гильз_ВК")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol11 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка гильз_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Грунтовка гильз_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                this.SymbolMetPipes = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Металл для крепления_ОВ_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Металл для крепления_ОВ_ВК")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                this.SymbolMetDuct = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Крепеж для воздуховодов_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Крепеж для воздуховодов_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                this.SymbolClei = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Клей для труб_ОВ_ВК_ИТП")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Клей для труб_ОВ_ВК_ИТП")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol12 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Гофра с зондом для капилярной трубки_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Гофра с зондом для капилярной трубки_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol13 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Гофра для электрического кабеля_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Гофра для электрического кабеля_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol14 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Кабель 3-х жильный_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Кабель 3-х жильный_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                FamilySymbol familySymbol15 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Кабель 4-х жильный_ОВ")).FirstOrDefault<Element>() as FamilySymbol;
//                if (((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Кабель 4-х жильный_ОВ")).Count<Element>() == 0)
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                if (!(((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Мастика_ОВ")).FirstOrDefault<Element>() is FamilySymbol familySymbol16))
//                {
//                    TaskDialog.Show("Ошибка!", "Отсутствует семейство ASML_ОВ_ВК_Элементы2D или загружена старая версия, загрузите или обновите и повторите запуск программы.");
//                    return (Result)0;
//                }
//                List<GlueDuctDto> glueDuctDtos = new List<GlueDuctDto>();
//                new GlueDuct().Get(this.Doc, estParam, glueDuctDtos);
//                if (!estParam.CheckParametersReadOnly(this.Doc, estParam, this.SymbolClei))
//                    return (Result)0;
//                using (Transaction transaction = new Transaction(this.Doc))
//                {
//                    transaction.Start("Add");
//                    if (this.CheckParams)
//                        return (Result)0;
//                    if (((Element)familySymbol15).LookupParameter("ASML_Версия семейства").AsString() == "v4_2022.11.23")
//                    {
//                        TaskDialog.Show("Ошибка!", "Обновите семейство ASML_ОВ_ВК_Элементы2D, до версии не ниже v5_2022.11.28");
//                        return (Result)0;
//                    }
//                    IList<ElementId> list8 = (IList<ElementId>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilyInstance))).Where<Element>((Func<Element, bool>)(x => x.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString() == "ASML_ОВ_ВК_Элементы2D")).Select<Element, ElementId>((Func<Element, ElementId>)(y => y.Id)).ToList<ElementId>();
//                    if (list8.Count > 0)
//                        this.Doc.Delete((ICollection<ElementId>)list8);
//                    this.AddMetalAndGluePipesFitting(estParam);
//                    if (Path.GetFileName(this.Doc.PathName).Contains("_ИТП") | Path.GetFileName(this.Doc.PathName).Contains("_ТМ"))
//                        new Itp().AddNotModElemItp(this.Doc);
//                    ICollection<Element> collectorLavList = (ICollection<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(Level)).ToElements();
//                    this.LevMax = collectorLavList.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble() == collectorLavList.Max<Element>((Func<Element, double>)(x => x.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble())))).FirstOrDefault<Element>() as Level;
//                    FilteredElementCollector source3 = new FilteredElementCollector(this.Doc).OfClass(typeof(Duct));
//                    List<string> strings1 = ((IEnumerable<Element>)source3).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamFolder) != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamSmeta).AsValueString() + a.get_Parameter(estParam.GuidParamSmeta).HasValue.ToString() != "НетTrue")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(estParam.GuidParamFolder).AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple1 = this.parameters.CheckNull(in strings1, "СМ_Папка Работ", "воздуховоды");
//                    if (tuple1.Item2)
//                        return (Result)0;
//                    List<string> stringList2 = tuple1.Item1;
//                    List<string> strings2 = ((IEnumerable<Element>)source3).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamLevel) != null)).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(estParam.GuidParamLevel).AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple2 = this.parameters.CheckNull(in strings2, "СМ_Этаж", "воздуховоды");
//                    if (tuple2.Item2)
//                        return (Result)0;
//                    strings2 = tuple2.Item1;
//                    List<string> strings3 = ((IEnumerable<Element>)source3).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamSection) != null)).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamSection).AsString() != "")).Select<Element, string>((Func<Element, string>)(a => a.get_Parameter(estParam.GuidParamSection).AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple3 = this.parameters.CheckNull(in strings3, "СМ_Секция", "воздуховоды");
//                    if (tuple3.Item2)
//                        return (Result)0;
//                    strings3 = tuple3.Item1;
//                    foreach (string str1 in stringList2)
//                    {
//                        string PapkaRabot = str1;
//                        List<Element> list9 = ((IEnumerable<Element>)source3).Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamFolder).AsString() == PapkaRabot)).ToList<Element>();
//                        foreach (string str2 in strings2)
//                        {
//                            string Level = str2;
//                            IList<Element> list10 = (IList<Element>)list9.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamLevel).AsString() == Level)).ToList<Element>();
//                            foreach (string str3 in strings3)
//                            {
//                                string Section = str3;
//                                List<Element> list11 = list10.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(estParam.GuidParamSection).AsString() == Section)).ToList<Element>();
//                                if (list11.Count<Element>() != 0)
//                                {
//                                    double num = this.Calc.GetMetallSumIsoDuct((IList<Element>)list11);
//                                    if (num < 0.001)
//                                        num = 0.001;
//                                    this.Ij += 0.25;
//                                    try
//                                    {
//                                        this.SymbolMetDuct.Activate();
//                                    }
//                                    catch
//                                    {
//                                        TaskDialog.Show("Ошибка!", "Отсутствует тип Крепеж для воздуховодов_ОВ в семействе ASML_ОВ_ВК_Элементы2D, обновите семейство до версии не ниже v4_2022.11.23, расчет остановлен!");
//                                        return (Result)0;
//                                    }
//                                    FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), this.SymbolMetDuct, this.LevMax, (StructuralType)0);
//                                    ((Element)familyInstance)[estParam.GuidParamCount].Set(num);
//                                    ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                    ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                    ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                    try
//                                    {
//                                        ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(list11.First<Element>()[estParam.GuidParamDiscipline].AsString());
//                                    }
//                                    catch
//                                    {
//                                    }
//                                    try
//                                    {
//                                        ((Element)familyInstance)[estParam.GuidParamChapter].Set(list11.First<Element>()[estParam.GuidParamChapter].AsString());
//                                    }
//                                    catch
//                                    {
//                                    }
//                                }
//                            }
//                        }
//                    }
//                    FilteredElementCollector source4 = new FilteredElementCollector(this.Doc).OfClass(typeof(Pipe));
//                    List<string> strings4 = ((IEnumerable<Element>)source4).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder] != null)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSmeta].AsValueString() + a[estParam.GuidParamSmeta].HasValue.ToString() != "НетTrue")).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamFolder].AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple4 = this.parameters.CheckNull(in strings4, "СМ_Папка Работ", "труб");
//                    if (tuple4.Item2)
//                        return (Result)0;
//                    strings4 = tuple4.Item1;
//                    List<string> strings5 = ((IEnumerable<Element>)source4).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection] != null)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection].AsString() != "")).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamSection].AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple5 = this.parameters.CheckNull(in strings5, "СМ_Секция", "труб");
//                    if (tuple5.Item2)
//                        return (Result)0;
//                    strings5 = tuple5.Item1;
//                    List<string> strings6 = ((IEnumerable<Element>)source4).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel] != null)).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamLevel].AsString())).Distinct<string>().ToList<string>();
//                    (List<string>, bool) tuple6 = this.parameters.CheckNull(in strings6, "СМ_Этаж", "труб");
//                    if (tuple6.Item2)
//                        return (Result)0;
//                    strings6 = tuple6.Item1;
//                    double num1 = 0.0;
//                    int num2 = 0;
//                    if (Path.GetFileName(this.Doc.PathName).Contains("_ИТП") | Path.GetFileName(this.Doc.PathName).Contains("_ТМ"))
//                    {
//                        if (num2 == 0)
//                            TaskDialog.Show("Предупреждение!", "Расчет металла для труб производится не будет т.к. вы находитесь в модели ИТП(ТМ), металл в данной модели учитывается элементами категории каркас несущий!");
//                        num2 = 1;
//                    }
//                    foreach (string str4 in strings4)
//                    {
//                        string PapkaRabot = str4;
//                        IList<Element> list12 = (IList<Element>)((IEnumerable<Element>)source4).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == PapkaRabot)).ToList<Element>();
//                        string str5 = "";
//                        try
//                        {
//                            str5 = list12.First<Element>()[estParam.GuidParamDiscipline].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        string str6 = "";
//                        try
//                        {
//                            str6 = list12.First<Element>()[estParam.GuidParamChapter].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        foreach (string str7 in strings6)
//                        {
//                            string Level = str7;
//                            IList<Element> list13 = (IList<Element>)list12.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel].AsString() == Level)).ToList<Element>();
//                            foreach (string str8 in strings5)
//                            {
//                                string Section = str8;
//                                IList<Element> list14 = (IList<Element>)list13.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection].AsString() == Section)).ToList<Element>();
//                                if (list14.Count<Element>() != 0)
//                                {
//                                    ElementParameterFilter filterStringEquals1 = this._parameterFilter.GetFilterStringEquals(estParam.idParamFolder, PapkaRabot);
//                                    ElementParameterFilter filterStringEquals2 = this._parameterFilter.GetFilterStringEquals(estParam.idParamLevel, Level);
//                                    ElementParameterFilter filterStringEquals3 = this._parameterFilter.GetFilterStringEquals(estParam.idParamSection, Section);
//                                    IList<Element> list15 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe)).Name.ToString() == "ASML_ОВ_Трубка капиллярная Ballu")).ToList<Element>();
//                                    if (list15.Count<Element>() != 0)
//                                    {
//                                        double num3 = list15.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        try
//                                        {
//                                            familySymbol12.Activate();
//                                        }
//                                        catch
//                                        {
//                                            TaskDialog.Show("Ошибка!", "Отсутствует тип Гофра с зондом для капилярной трубки_ОВ в семействе ASML_ОВ_ВК_Элементы2D, обновите семейство до версии не ниже v4_2022.11.23, расчет остановлен!!!");
//                                            return (Result)0;
//                                        }
//                                        FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol12, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance)[estParam.GuidParamCount].Set(num3);
//                                        ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                    }
//                                    IList<Element> list16 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe)).Name.ToString().Contains("ASML_ОВ_Медь Сварка"))).ToList<Element>();
//                                    if (list16.Count<Element>() != 0 && flag)
//                                    {
//                                        double num4 = list16.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        try
//                                        {
//                                            familySymbol13.Activate();
//                                        }
//                                        catch
//                                        {
//                                            TaskDialog.Show("Ошибка!", "Отсутствует тип Гофра для электрического кабеля_ОВ в семействе ASML_ОВ_ВК_Элементы2D, обновите семейство до версии не ниже v4_2022.11.23, расчет остановлен!!!");
//                                            return (Result)0;
//                                        }
//                                        FamilyInstance familyInstance1 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol13, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance1)[estParam.GuidParamCount].Set(num4);
//                                        ((Element)familyInstance1)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance1)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance1)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance1)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance1)[estParam.GuidParamChapter].Set(str6);
//                                        this.Ij += 0.25;
//                                        try
//                                        {
//                                            familySymbol14.Activate();
//                                        }
//                                        catch
//                                        {
//                                            TaskDialog.Show("Ошибка!", "Отсутствует тип Кабель 3 - х жильный_ОВ в семействе ASML_ОВ_ВК_Элементы2D, обновите семейство до версии не ниже v4_2022.11.23, расчет остановлен!!!");
//                                            return (Result)0;
//                                        }
//                                        FamilyInstance familyInstance2 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol14, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance2)[estParam.GuidParamCount].Set(num4 / 2.0);
//                                        ((Element)familyInstance2)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance2)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance2)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance2)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance2)[estParam.GuidParamChapter].Set(str6);
//                                        this.Ij += 0.25;
//                                        try
//                                        {
//                                            familySymbol15.Activate();
//                                        }
//                                        catch
//                                        {
//                                            TaskDialog.Show("Ошибка!", "Отсутствует тип Кабель 4 - х жильный_ОВ в семействе ASML_ОВ_ВК_Элементы2D, обновите семейство до версии не ниже v4_2022.11.23, расчет остановлен!!!");
//                                            return (Result)0;
//                                        }
//                                        FamilyInstance familyInstance3 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol15, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance3)[estParam.GuidParamCount].Set(num4 / 2.0);
//                                        ((Element)familyInstance3)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance3)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance3)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance3)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance3)[estParam.GuidParamChapter].Set(str6);
//                                    }
//                                    IList<Element> list17 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe).PipeType).GetParameters("Комментарии к типоразмеру").FirstOrDefault<Parameter>().AsString() == "Цинол_ОВ_ВК_ИТП")).ToList<Element>();
//                                    IList<Element> list18 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe).PipeType).GetParameters("Комментарии к типоразмеру").FirstOrDefault<Parameter>().AsString() == "Грунтовка_Краска_ВК")).ToList<Element>();
//                                    IList<Element> list19 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe).PipeType).GetParameters("Комментарии к типоразмеру").FirstOrDefault<Parameter>().AsString() == "Грунтовка_Краска_ОВ")).ToList<Element>();
//                                    IList<Element> list20 = (IList<Element>)list14.Where<Element>((Func<Element, bool>)(a => ((Element)(a as Pipe).PipeType).GetParameters("Комментарии к типоразмеру").FirstOrDefault<Parameter>().AsString() == "Краска_ИТП")).ToList<Element>();
//                                    if (list17.Count<Element>() != 0)
//                                    {
//                                        double num5 = list17.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 / 0.3 * 0.1 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        familySymbol1.Activate();
//                                        FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol1, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance)[estParam.GuidParamCount].Set(num5);
//                                        ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                    }
//                                    if (list18.Count<Element>() != 0)
//                                    {
//                                        IList<Element> elementList1 = (IList<Element>)new List<Element>();
//                                        IList<Element> list21 = (IList<Element>)list18.Where<Element>((Func<Element, bool>)(a => (a as MEPCurve).MEPSystem != null)).Where<Element>((Func<Element, bool>)(a => this.Doc.GetElement(((Element)(a as MEPCurve).MEPSystem).GetTypeId())[(BuiltInParameter) - 1010105].AsString() != "Краска_АУВПТ")).Concat<Element>(list18.Where<Element>((Func<Element, bool>)(a => (a as MEPCurve).MEPSystem == null))).ToList<Element>();
//                                        if (list21.Count<Element>() != 0)
//                                        {
//                                            double num6 = list21.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                            this.Ij += 0.25;
//                                            familySymbol2.Activate();
//                                            FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol2, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance)[estParam.GuidParamCount].Set(num6);
//                                            ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                            ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                            ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                            ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                            ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                        }
//                                        IList<Element> elementList2 = (IList<Element>)new List<Element>();
//                                        IList<Element> list22 = (IList<Element>)list18.Where<Element>((Func<Element, bool>)(a => (a as MEPCurve).MEPSystem != null)).Where<Element>((Func<Element, bool>)(a => this.Doc.GetElement(((Element)(a as MEPCurve).MEPSystem).GetTypeId())[(BuiltInParameter) - 1010105].AsString() == "Краска_АУВПТ")).ToList<Element>();
//                                        if (list22.Count<Element>() != 0)
//                                        {
//                                            double num7 = list22.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                            double num8 = list22.Where<Element>((Func<Element, bool>)(a => !((Element)(a as MEPCurve).MEPSystem)[(BuiltInParameter) - 1150468].AsString().Contains("В23"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                            if (num8 > 0.0)
//                                            {
//                                                this.Ij += 0.25;
//                                                familySymbol3.Activate();
//                                                FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol3, this.LevMax, (StructuralType)0);
//                                                ((Element)familyInstance)[estParam.GuidParamCount].Set(num8 * 0.9);
//                                                ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                                ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                                ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                                ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                                ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                            }
//                                            double num9 = list22.Where<Element>((Func<Element, bool>)(a => ((Element)(a as MEPCurve).MEPSystem)[(BuiltInParameter) - 1150468].AsString().Contains("В23"))).Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                            if (num9 > 0.0)
//                                            {
//                                                this.Ij += 0.25;
//                                                familySymbol5.Activate();
//                                                FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol5, this.LevMax, (StructuralType)0);
//                                                ((Element)familyInstance)[estParam.GuidParamCount].Set(num9 * 0.9);
//                                                ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                                ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                                ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                                ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                                ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                            }
//                                            this.Ij += 0.25;
//                                            familySymbol4.Activate();
//                                            FamilyInstance familyInstance4 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol4, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance4)[estParam.GuidParamCount].Set(num7 * 0.1);
//                                            ((Element)familyInstance4)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                            ((Element)familyInstance4)[estParam.GuidParamSection].Set(Section);
//                                            ((Element)familyInstance4)[estParam.GuidParamLevel].Set(Level);
//                                            ((Element)familyInstance4)[estParam.GuidParamDiscipline].Set(str5);
//                                            ((Element)familyInstance4)[estParam.GuidParamChapter].Set(str6);
//                                        }
//                                        double num10 = list18.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        familySymbol8.Activate();
//                                        FamilyInstance familyInstance5 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol8, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance5)[estParam.GuidParamCount].Set(num10);
//                                        ((Element)familyInstance5)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance5)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance5)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance5)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance5)[estParam.GuidParamChapter].Set(str6);
//                                        elementList1 = (IList<Element>)null;
//                                        elementList2 = (IList<Element>)null;
//                                    }
//                                    if (list19.Count<Element>() != 0)
//                                    {
//                                        double num11 = list19.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        familySymbol6.Activate();
//                                        FamilyInstance familyInstance6 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol6, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance6)[estParam.GuidParamCount].Set(num11);
//                                        ((Element)familyInstance6)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance6)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance6)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance6)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance6)[estParam.GuidParamChapter].Set(str6);
//                                        this.Ij += 0.25;
//                                        familySymbol9.Activate();
//                                        FamilyInstance familyInstance7 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol9, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance7)[estParam.GuidParamCount].Set(num11);
//                                        ((Element)familyInstance7)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance7)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance7)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance7)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance7)[estParam.GuidParamChapter].Set(str6);
//                                    }
//                                    if (list20.Count<Element>() != 0)
//                                    {
//                                        double num12 = list20.Select<Element, double>((Func<Element, double>)(a => a.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble() / (1250.0 / 381.0) * 1.2 * a[(BuiltInParameter) - 1140238].AsDouble() * 3.14 / (1250.0 / 381.0))).Sum();
//                                        this.Ij += 0.25;
//                                        familySymbol7.Activate();
//                                        FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol7, this.LevMax, (StructuralType)0);
//                                        ((Element)familyInstance)[estParam.GuidParamCount].Set(num12);
//                                        ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                        ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                        ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                        ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                        ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                    }
//                                    double num13 = 0.0;
//                                    if (num2 == 1)
//                                    {
//                                        try
//                                        {
//                                            num13 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilyInstance))).Where<Element>((Func<Element, bool>)(x => x.Name == "60x60x6" | x.Name == "L 50x5" | x.Name == "L 63x6" | x.Name == "6.5У" | x.Name == "10У" | x.Name == "16П")).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == PapkaRabot)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection].AsString() == Section)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel].AsString() == Level)).Select<Element, double>((Func<Element, double>)(z => z.LookupParameter("ASML_Площадь").AsDouble() * 0.09290304)).Sum();
//                                        }
//                                        catch
//                                        {
//                                            num1 = 1.0;
//                                        }
//                                        if (num13 != 0.0)
//                                        {
//                                            this.Ij += 0.25;
//                                            familySymbol7.Activate();
//                                            FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol7, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance)[estParam.GuidParamCount].Set(num13);
//                                            ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                            ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                            ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                            ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                            ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                        }
//                                    }
//                                    try
//                                    {
//                                        ElementParameterFilter filterStringEquals4 = this._parameterFilter.GetFilterStringEquals(id1, "ASML_ВК_Гильза");
//                                        IList<double> list23 = (IList<double>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilyInstance)).WherePasses((ElementFilter)filterStringEquals4).WherePasses((ElementFilter)filterStringEquals1).WherePasses((ElementFilter)filterStringEquals2).WherePasses((ElementFilter)filterStringEquals3)).Select<Element, double>((Func<Element, double>)(y => y[estParam.GuidParamDnar].AsDouble() * 304.8 * 3.14 / 1000.0 * y.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 * 2.0 / 1000.0)).ToList<double>();
//                                        ElementParameterFilter filterStringEquals5 = this._parameterFilter.GetFilterStringEquals(id1, "ASML_ОВ_Гильза");
//                                        IList<double> list24 = (IList<double>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilyInstance)).WherePasses((ElementFilter)filterStringEquals5).WherePasses((ElementFilter)filterStringEquals1).WherePasses((ElementFilter)filterStringEquals2).WherePasses((ElementFilter)filterStringEquals3)).Select<Element, double>((Func<Element, double>)(y => y[estParam.GuidParamDnar].AsDouble() * 304.8 * 3.14 / 1000.0 * y.get_Parameter(estParam.GuidParamSizeLanthFitt).AsDouble() * 304.8 * 2.0 / 1000.0)).ToList<double>();
//                                        double num14 = list23.Sum();
//                                        double num15 = list24.Sum();
//                                        if (num14 != 0.0)
//                                        {
//                                            this.Ij += 0.25;
//                                            familySymbol10.Activate();
//                                            FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol10, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance)[estParam.GuidParamCount].Set(num14 * 1.2);
//                                            try
//                                            {
//                                                ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                                ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                                ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                                ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                                ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                            }
//                                            catch
//                                            {
//                                                TaskDialog.Show("Предупреждение!", "Проверьте заполнение СМ_Папка работ в семействе: " + ((Element)familyInstance).Name.ToString());
//                                                return (Result)0;
//                                            }
//                                        }
//                                        if (num15 != 0.0)
//                                        {
//                                            this.Ij += 0.25;
//                                            familySymbol11.Activate();
//                                            FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol11, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance)[estParam.GuidParamCount].Set(num15 * 1.2);
//                                            ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                            ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                            ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                            ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                            ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                        }
//                                    }
//                                    catch
//                                    {
//                                        TaskDialog.Show("Ошибка!", "Отсутствует параметр ASML_Диаметр наружный в гильзах, обновите семейство гильз до версии не ниже v1_2022.10.25, расчет остановлен!");
//                                        return (Result)0;
//                                    }
//                                    if (num2 == 0)
//                                    {
//                                        ElementParameterFilter filterDoubleNotEquals = this._parameterFilter.GetFilterDoubleNotEquals(parameterId2, 0.0);
//                                        ElementParameterFilter filterStringNotEquals = this._parameterFilter.GetFilterStringNotEquals(id2, "Крепления_НЕТ");
//                                        IList<Element> list25 = (IList<Element>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(Pipe)).WherePasses((ElementFilter)filterStringEquals1).WherePasses((ElementFilter)filterStringEquals2).WherePasses((ElementFilter)filterStringEquals3).WherePasses((ElementFilter)filterDoubleNotEquals).WherePasses((ElementFilter)filterStringNotEquals)).ToList<Element>();
//                                        ElementParameterFilter filterDoubleEquals = this._parameterFilter.GetFilterDoubleEquals(parameterId2, 0.0);
//                                        IList<Element> list26 = (IList<Element>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(Pipe)).WherePasses((ElementFilter)filterStringEquals1).WherePasses((ElementFilter)filterStringEquals2).WherePasses((ElementFilter)filterStringEquals3).WherePasses((ElementFilter)filterDoubleEquals).WherePasses((ElementFilter)filterStringNotEquals)).ToList<Element>();
//                                        if (list25.Count<Element>() > 0 | list26.Count<Element>() > 0)
//                                        {
//                                            this.SymbolMetPipes.Activate();
//                                            double num16 = this.Calc.GetPipesMetallSum(list25) + this.Calc.GetPipesMetallSumIso(list26);
//                                            if (num16 != 0.0)
//                                            {
//                                                if (num16 < 0.001)
//                                                    num16 = 0.001;
//                                                this.Ij += 0.25;
//                                                FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), this.SymbolMetPipes, this.LevMax, (StructuralType)0);
//                                                ((Element)familyInstance)[estParam.GuidParamCount].Set(num16);
//                                                ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                                ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                                ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                                ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                                ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                            }
//                                        }
//                                    }
//                                    ElementParameterFilter filterStringNotEquals1 = this._parameterFilter.GetFilterStringNotEquals(parameterId1, "ASML_Гофра");
//                                    IList<ElementId> list27 = (IList<ElementId>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(PipeInsulation)).WherePasses((ElementFilter)filterStringEquals1).WherePasses((ElementFilter)filterStringEquals2).WherePasses((ElementFilter)filterStringEquals3).WherePasses((ElementFilter)filterStringNotEquals1)).Select<Element, ElementId>((Func<Element, ElementId>)(a => (a as InsulationLiningBase).HostElementId)).ToList<ElementId>();
//                                    IList<Element> elementList = (IList<Element>)new List<Element>();
//                                    IList<Element> list28;
//                                    try
//                                    {
//                                        list28 = (IList<Element>)list27.Where<ElementId>((Func<ElementId, bool>)(s => this.Doc.GetElement(s).Category.Name == "Трубы")).Select<ElementId, Element>((Func<ElementId, Element>)(a => this.Doc.GetElement(a))).ToList<Element>();
//                                    }
//                                    catch
//                                    {
//                                        TaskDialog.Show("Предупреждение!", "В модели присутствуют экземпляры изоляции трубопровода без основы. Проверьте и удалите изоляцию без основы плагином по проверке изоляции.");
//                                        return (Result)0;
//                                    }
//                                    double num17 = 0.0;
//                                    if (list27.Count<ElementId>() > 0)
//                                    {
//                                        double cleiPipe = this.Calc.GetCleiPipe((IList<Element>)list28.Where<Element>((Func<Element, bool>)(a => this._insulationPipeNotClei.Contains(a[(BuiltInParameter) - 1150430].AsString()))).ToList<Element>());
//                                        this.SymbolClei.Activate();
//                                        if (cleiPipe != 0.0)
//                                        {
//                                            this.Ij += 0.25;
//                                            FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), this.SymbolClei, this.LevMax, (StructuralType)0);
//                                            ((Element)familyInstance)[estParam.GuidParamCount].Set(cleiPipe);
//                                            ((Element)familyInstance)[estParam.GuidParamFolder].Set(PapkaRabot);
//                                            ((Element)familyInstance)[estParam.GuidParamSection].Set(Section);
//                                            ((Element)familyInstance)[estParam.GuidParamLevel].Set(Level);
//                                            ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(str5);
//                                            ((Element)familyInstance)[estParam.GuidParamChapter].Set(str6);
//                                        }
//                                    }
//                                    num17 = 0.0;
//                                }
//                            }
//                        }
//                    }
//                    if (num1 == 1.0 && Path.GetFileName(this.Doc.PathName).Contains("_ИТП") | Path.GetFileName(this.Doc.PathName).Contains("_ТМ") && MessageBox.Show("В проект отсутствуют элементы металлоконструкций или используются старые семейства, необходимо заменить на новые, см.список типов ниже:\n60x60x6\nL 50x5\nL 63x6\n6.5У\n10У\n16П!\nПродолжить расчет без краски металлоконструкций?", "Предупреждение!", MessageBoxButton.YesNo) == MessageBoxResult.No)
//                        return (Result)0;
//                    IList<Element> elementList3 = (IList<Element>)new List<Element>();
//                    IList<Element> list29 = (IList<Element>)new FilteredElementCollector(this.Doc).WhereElementIsNotElementType().ToElements().Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder] != null)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == "В1 (ВОДОМЕРНЫЙ УЗЕЛ)")).ToList<Element>();
//                    if (list29.Count<Element>() != 0)
//                    {
//                        string str9 = "";
//                        try
//                        {
//                            str9 = list29.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamDiscipline].AsString().Length > 1)).First<Element>()[estParam.GuidParamDiscipline].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        string str10 = "";
//                        try
//                        {
//                            str10 = list29.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamChapter].AsString().Length > 1)).First<Element>()[estParam.GuidParamChapter].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        string str11 = "";
//                        try
//                        {
//                            str11 = list29.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection].AsString().Length > 1)).First<Element>()[estParam.GuidParamSection].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        string str12 = "";
//                        try
//                        {
//                            str12 = list29.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel].AsString().Length > 1)).First<Element>()[estParam.GuidParamLevel].AsString();
//                        }
//                        catch
//                        {
//                        }
//                        FamilySymbol familySymbol17 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "ОпораКНС_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                        this.Ij += 0.25;
//                        familySymbol17.Activate();
//                        FamilyInstance familyInstance8 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol17, this.LevMax, (StructuralType)0);
//                        ((Element)familyInstance8)[estParam.GuidParamFolder].Set("В1 (ВОДОМЕРНЫЙ УЗЕЛ)");
//                        ((Element)familyInstance8)[estParam.GuidParamDiscipline].Set(str9);
//                        ((Element)familyInstance8)[estParam.GuidParamChapter].Set(str10);
//                        ((Element)familyInstance8)[estParam.GuidParamSection].Set(str11);
//                        ((Element)familyInstance8)[estParam.GuidParamLevel].Set(str12);
//                        FamilySymbol familySymbol18 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "ОпораОП_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                        this.Ij += 0.25;
//                        familySymbol18.Activate();
//                        FamilyInstance familyInstance9 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol18, this.LevMax, (StructuralType)0);
//                        ((Element)familyInstance9)[estParam.GuidParamFolder].Set("В1 (ВОДОМЕРНЫЙ УЗЕЛ)");
//                        ((Element)familyInstance9)[estParam.GuidParamDiscipline].Set(str9);
//                        ((Element)familyInstance9)[estParam.GuidParamChapter].Set(str10);
//                        ((Element)familyInstance9)[estParam.GuidParamSection].Set(str11);
//                        ((Element)familyInstance9)[estParam.GuidParamLevel].Set(str12);
//                        FamilySymbol familySymbol19 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Узел усиления_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                        this.Ij += 0.25;
//                        familySymbol19.Activate();
//                        FamilyInstance familyInstance10 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol19, this.LevMax, (StructuralType)0);
//                        ((Element)familyInstance10)[estParam.GuidParamFolder].Set("В1 (ВОДОМЕРНЫЙ УЗЕЛ)");
//                        ((Element)familyInstance10)[estParam.GuidParamDiscipline].Set(str9);
//                        ((Element)familyInstance10)[estParam.GuidParamChapter].Set(str10);
//                        ((Element)familyInstance10)[estParam.GuidParamSection].Set(str11);
//                        ((Element)familyInstance10)[estParam.GuidParamLevel].Set(str12);
//                        FamilySymbol familySymbol20 = ((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfClass(typeof(FamilySymbol))).Where<Element>((Func<Element, bool>)(x => x.Name == "Упор бетонный_ВК")).FirstOrDefault<Element>() as FamilySymbol;
//                        this.Ij += 0.25;
//                        familySymbol20.Activate();
//                        FamilyInstance familyInstance11 = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol20, this.LevMax, (StructuralType)0);
//                        ((Element)familyInstance11)[estParam.GuidParamFolder].Set("В1 (ВОДОМЕРНЫЙ УЗЕЛ)");
//                        ((Element)familyInstance11)[estParam.GuidParamDiscipline].Set(str9);
//                        ((Element)familyInstance11)[estParam.GuidParamChapter].Set(str10);
//                        ((Element)familyInstance11)[estParam.GuidParamSection].Set(str11);
//                        ((Element)familyInstance11)[estParam.GuidParamLevel].Set(str12);
//                        TaskDialog.Show("Предупреждение!", "В проект добавлены не моделируемые элементов для водомерного узла, не забудьте заполнить сметные параметры!");
//                    }
//                    foreach (GlueDuctDto glueDuctDto in glueDuctDtos)
//                    {
//                        this.Ij += 0.25;
//                        familySymbol16.Activate();
//                        FamilyInstance familyInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), familySymbol16, this.LevMax, (StructuralType)0);
//                        ((Element)familyInstance)[estParam.GuidParamCount].Set(glueDuctDto.Quantity);
//                        ((Element)familyInstance)[estParam.GuidParamFolder].Set(glueDuctDto.FolderJob);
//                        ((Element)familyInstance)[estParam.GuidParamSection].Set(glueDuctDto.Section);
//                        ((Element)familyInstance)[estParam.GuidParamLevel].Set(glueDuctDto.Levels);
//                        ((Element)familyInstance)[estParam.GuidParamDiscipline].Set(glueDuctDto.Discipline);
//                        ((Element)familyInstance)[estParam.GuidParamChapter].Set(glueDuctDto.Chapter);
//                    }
//                    transaction.Commit();
//                }
//                stopwatch.Stop();
//                TaskDialog.Show(stopwatch.Elapsed.ToString(), "Расчет выполнен, семейства добавлены в проект!");
//                return (Result)0;
//            }

//            public void SetSmPameters(Specification.EstimateParameters estParam)
//            {
//                ((Element)this.ElementInstance)[estParam.GuidParamCount].Set(this.Quantity);
//                ((Element)this.ElementInstance)[estParam.GuidParamFolder].Set(this.FolderJob);
//                ((Element)this.ElementInstance)[estParam.GuidParamSection].Set(this.Section);
//                ((Element)this.ElementInstance)[estParam.GuidParamLevel].Set(this.Levels);
//                ((Element)this.ElementInstance)[estParam.GuidParamDiscipline].Set(this.Discipline);
//                ((Element)this.ElementInstance)[estParam.GuidParamChapter].Set(this.Chapter);
//            }

//            public void AddMetalAndGluePipesFitting(Specification.EstimateParameters estParam)
//            {
//                ParameterFilters parameterFilters = new ParameterFilters();
//                IList<Element> elementList = (IList<Element>)new List<Element>();
//                IList<Element> list1 = (IList<Element>)((IEnumerable<Element>)new FilteredElementCollector(this.Doc).OfCategory((BuiltInCategory) - 2008049).WhereElementIsNotElementType()).Where<Element>((Func<Element, bool>)(a => this.Doc.GetElement(a.GetTypeId())[(BuiltInParameter) - 1010105].AsString() == "Труба")).ToList<Element>();
//                if (estParam.CheckParameters(estParam, list1))
//                {
//                    this.CheckParams = true;
//                }
//                else
//                {
//                    if (list1.Count<Element>() == 0)
//                        return;
//                    IList<string> list2 = (IList<string>)list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder] != null)).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamFolder].AsString())).Distinct<string>().ToList<string>();
//                    IList<string> list3 = (IList<string>)list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection] != null)).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamSection].AsString())).Distinct<string>().ToList<string>();
//                    IList<string> list4 = (IList<string>)list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel] != null)).Select<Element, string>((Func<Element, string>)(a => a[estParam.GuidParamLevel].AsString())).Distinct<string>().ToList<string>();
//                    foreach (string str1 in (IEnumerable<string>)list2)
//                    {
//                        string folderJob = str1;
//                        this.FolderJob = folderJob;
//                        this.Discipline = list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == folderJob)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamDiscipline].AsString().Length > 1)).FirstOrDefault<Element>()[estParam.GuidParamDiscipline].AsString();
//                        this.Chapter = list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == folderJob)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamChapter].AsString().Length > 1)).FirstOrDefault<Element>()[estParam.GuidParamChapter].AsString();
//                        foreach (string str2 in (IEnumerable<string>)list3)
//                        {
//                            string section = str2;
//                            this.Section = section;
//                            foreach (string str3 in (IEnumerable<string>)list4)
//                            {
//                                string level = str3;
//                                this.Levels = level;
//                                IList<Element> list5 = (IList<Element>)list1.Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamFolder].AsString() == folderJob)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamSection].AsString() == section)).Where<Element>((Func<Element, bool>)(a => a[estParam.GuidParamLevel].AsString() == level)).ToList<Element>();
//                                IList<Element> list6 = (IList<Element>)list5.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() == 0.0)).ToList<Element>();
//                                IList<Element> list7 = (IList<Element>)list5.Where<Element>((Func<Element, bool>)(a => a.get_Parameter(BuiltInParameter.RBS_REFERENCE_INSULATION_THICKNESS).AsDouble() != 0.0)).ToList<Element>();
//                                double num1 = 0.0;
//                                double num2 = this.Calc.GetFittMetallSumIso(this.Doc, estParam, list6) + this.Calc.GetFittMetallSum(this.Doc, estParam, list7);
//                                if (num2 > 0.0)
//                                {
//                                    if (num2 < 0.001)
//                                        num2 = 0.001;
//                                    this.Ij += 0.25;
//                                    this.SymbolMetPipes.Activate();
//                                    this.ElementInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), this.SymbolMetPipes, this.LevMax, (StructuralType)0);
//                                    this.Quantity = num2;
//                                    this.SetSmPameters(estParam);
//                                }
//                                double cleiFitt = this.Calc.GetCleiFitt(estParam, (IList<Element>)list7.Where<Element>((Func<Element, bool>)(a => this._insulationPipeNotClei.Contains(a[(BuiltInParameter) - 1150430].AsString()))).ToList<Element>());
//                                if (cleiFitt > 0.0)
//                                {
//                                    if (num2 < 0.001)
//                                        num1 = 0.001;
//                                    this.Quantity = cleiFitt;
//                                    this.SymbolClei.Activate();
//                                    this.Ij += 0.25;
//                                    this.ElementInstance = this.Doc.Create.NewFamilyInstance(new XYZ(this.Ij, 0.0, 300.0), this.SymbolClei, this.LevMax, (StructuralType)0);
//                                    this.SetSmPameters(estParam);
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//        }
//    }
//}
