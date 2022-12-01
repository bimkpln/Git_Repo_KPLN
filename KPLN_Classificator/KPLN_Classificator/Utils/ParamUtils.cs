using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Classificator.ApplicationConfig;
using Autodesk.Revit.DB.Architecture;

namespace KPLN_Classificator
{
    public class ParamUtils
    {
        public List<Element> fullSuccessElems = new List<Element>();
        public List<Element> notFullSuccessElems = new List<Element>();
        private static bool debug;

        public ParamUtils(bool debugMode)
        {
            debug = debugMode;
        }

        public static bool nameChecker(string nameClafi, string nameElem)
        {
            string[] arrayClafiAnd = nameClafi.ToLower().Split(',');
            int index = arrayClafiAnd.Length;

            if (index == 1)
            {
                if (arrayClafiAnd.First().StartsWith("!"))
                {
                    return !nameElem.ToLower().Contains(arrayClafiAnd.First().Replace("!","")) && !nameElem.Contains("!");
                }

                if (arrayClafiAnd.First().Contains('|'))
                {
                    bool check = false;
                    string[] arrayClafiOr = arrayClafiAnd.First().Split('|');
                    for (int j = 0; j < arrayClafiOr.Length; j++)
                    {
                        if (nameElem.ToLower().Contains(arrayClafiOr[j]))
                        {
                            check = true;
                            break;
                        }
                    }
                    return check;
                }

                return nameElem.ToLower().Contains(arrayClafiAnd.First());
            }
            else if (index > 1)
            {
                for (int i = 0; i < index; i++)
                {
                    if (arrayClafiAnd[i].StartsWith("!"))
                    {
                        if (!nameElem.ToLower().Contains(arrayClafiAnd[i].Replace("!", "")) && !nameElem.Contains("!"))
                        {
                            continue;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else if (arrayClafiAnd[i].Contains('|'))
                    {
                        bool check = false;
                        string[] arrayClafiOr = arrayClafiAnd[i].Split('|');
                        for (int j = 0; j < arrayClafiOr.Length; j++)
                        {
                            if (nameElem.ToLower().Contains(arrayClafiOr[j]))
                            {
                                check = true;
                                break;
                            }
                        }
                        if (check) continue;
                        else return false;
                    }
                    else if (nameElem.ToLower().Contains(arrayClafiAnd[i])) continue;
                    else return false;
                }
                return true;
            }
            return false;
        }

        private static bool setParam(Element elem, string paramName, string value, out string valueForAssigned)
        {
            valueForAssigned = null;
            bool rsl = false;
            string newValue = value;
            if (value.Contains("[") && value.Contains("]"))
            {
                char[] valueArray = value.ToCharArray();
                Dictionary<string, string> foundParamsAndTheirValues = new Dictionary<string, string>();
                for (int i = 0; i < valueArray.Length; i++)
                {
                    if (valueArray[i] == '[')
                    {
                        for (int j = i; j < valueArray.Length; j++)
                        {
                            if (valueArray[j] == ']')
                            {
                                string foundParamName = value.Substring(i, j - i + 1);
                                string foundParamNameForGettingValue = foundParamName.Replace("]", "").Replace("[", "");
                                string valueOfParam = null;
                                if (foundParamName.Contains("*"))
                                {
                                    valueOfParam = getValueStringOfAllParams(elem, foundParamNameForGettingValue.Split('*')[0]);
                                    if (valueOfParam == null) return rsl;
                                    try
                                    {
                                        valueOfParam = multipleSourseParamOnMultiplier(valueOfParam, foundParamNameForGettingValue);
                                    }
                                    catch (Exception)
                                    {
                                        output.PrintDebug(string.Format("Значение параметра: \"{0}\" в классификаторе содержит операцию умножения (*), которое не было выполнено. Проверьте корректность заполнения конфигурационного файла. Значение не вписано в параметр: \"{1}\".", 
                                            foundParamName, paramName), Output.OutputMessageType.Warning, debug);
                                        return rsl;
                                    }
                                }
                                else
                                {
                                    valueOfParam = getValueStringOfAllParams(elem, foundParamNameForGettingValue);
                                    if (valueOfParam == null) valueOfParam = "";
                                }
                                if (!foundParamsAndTheirValues.ContainsKey(foundParamName))
                                {
                                    foundParamsAndTheirValues.Add(foundParamName, valueOfParam);
                                }
                                break;
                            }
                        }
                    }
                }
                foreach (string item in foundParamsAndTheirValues.Keys)
                {
                    string itemValue = foundParamsAndTheirValues[item];
                    if (itemValue == null || itemValue.Length == 0)
                    {
                        output.PrintDebug(string.Format("Не заполнено значение параметра: \"{0}\" у элемента: {1} с id: {2}. Значение не вписано в параметр: \"{3}\".", item, elem.Name, elem.Id, paramName), Output.OutputMessageType.Warning, debug);
                        return rsl;
                    }
                    newValue = newValue.Replace(item, itemValue);
                }
            }
            Parameter param = elem.LookupParameter(paramName);
            if (param == null)
            {
                output.PrintDebug(string.Format("В элементе: \"{0}\" с id: {1} не найден параметр: \"{2}\".", elem.Name, elem.Id, paramName), Output.OutputMessageType.Warning, debug);
                return rsl;
            }
            try
            {
                if (param.StorageType == StorageType.String)
                {
                    param.Set(newValue);
                }
                else if (param.StorageType == StorageType.Double)
                {
                    param.Set(double.Parse(newValue));
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    param.Set(int.Parse(newValue));
                }
                valueForAssigned = newValue;
                rsl = true;
            }
            catch (Exception)
            {
                output.PrintDebug(string.Format("Не удалось присвоить значение \"{0}\" параметру: \"{1}\" с типом данных: {2}. Элемент: {3} с id: {4}", 
                    newValue, paramName, param.StorageType.ToString(), elem.Name, elem.Id), Output.OutputMessageType.Warning, debug);
            }
            return rsl;
        }

        private void setClassificator(Classificator classificator, InfosStorage storage, Element elem)
        {
            bool paramChecker;
            List<string> assignedValues = new List<string>();

            if (classificator.paramsValues.Count > storage.instanseParams.Count)
            {
                output.PrintDebug(string.Format("Значение параметра: \"{0}\" в элементе: \"{1}\" за пределами диапазона возможных значений. Присвоение данного параметра не будет выполнено."
                    , classificator.paramsValues[classificator.paramsValues.Count - 1]
                    , classificator.FamilyName)
                    , Output.OutputMessageType.Warning, debug);
            }
            for (int i = 0; i < Math.Min(classificator.paramsValues.Count, storage.instanseParams.Count); i++)
            {
                if (classificator.paramsValues[i].Length == 0) continue;
                paramChecker = setParam(elem, storage.instanseParams[i], classificator.paramsValues[i], out string valueForAssigned);
                if (paramChecker)
                {
                    assignedValues.Add(valueForAssigned);
                }
            }
            if (assignedValues.Count == Math.Min(classificator.paramsValues.Where(i => i.Length > 0).Count(), storage.instanseParams.Where(i => i.Length > 0).Count()))
            {
                fullSuccessElems.Add(elem);
            }
            else
            {
                notFullSuccessElems.Add(elem);
            }
            output.PrintDebug(string.Format("Были присвоены значения: {0}", string.Join("; ", assignedValues)), Output.OutputMessageType.System_OK, debug);
        }

        public static string getValueStringOfAllParams(Element elem, string paramName)
        {
            string paramValue = null;
            Parameter paramObject = elem.LookupParameter(paramName);
            if (paramObject == null)
            {
                try
                {
                    paramObject = elem.Document.GetElement(elem.GetTypeId()).LookupParameter(paramName);
                }
                catch (Exception)
                { }
            }
            if (paramObject == null)
            {
                output.PrintDebug(string.Format("В элементе: \"{0}\" c id: {1} не найден параметр: \"{2}\"", elem.Name, elem.Id, paramName), Output.OutputMessageType.Warning, debug);
                return paramValue;
            }
            try
            {
                if (paramObject.StorageType == StorageType.String)
                {
                    paramValue = paramObject.AsString().ToString();
                }
                else if (paramObject.StorageType == StorageType.Double)
                {
                    paramValue = paramObject.AsDouble().ToString();
                }
                else if (paramObject.StorageType == StorageType.Integer)
                {
                    paramValue = paramObject.AsInteger().ToString();
                }
                else
                {
                    output.PrintDebug("Не удалось определить тип параметра: " + paramName, Output.OutputMessageType.Error, debug);
                }
            }
            catch (Exception) { }
            return paramValue;
        }

        private static string multipleSourseParamOnMultiplier(string valueOfParam, string foundParamNameForGettingValue) 
        {
            double paramValue = double.Parse(valueOfParam);
            double multiplier = double.Parse(foundParamNameForGettingValue.Split('*')[1].Split('D')[0].Replace(".", ","));
            int digits = 0;
            if (foundParamNameForGettingValue.Contains("D"))
            {
                int.TryParse(foundParamNameForGettingValue.Split('D')[1], out digits);
            }
            double result = paramValue * multiplier;
            return foundParamNameForGettingValue.Contains("D") ? Math.Round(result, digits == 0 ? 1 : digits).ToString() : Math.Round(result).ToString();
        }

        public bool startClassification(List<Element> constrs, InfosStorage storage, Document doc)
        {
            if (constrs == null || constrs.Count == 0)
            {
                output.PrintInfo("Не удалось получить элементы для заполнения классификатора! " +
                    "Проверь выборку (можно выбирать только моделируемые элементы), либо запусти на весь проект!", Output.OutputMessageType.Error);
                return false;
            }
            foreach (Classificator classificator in storage.classificator)
            {
                output.PrintDebug(string.Format("{0} - {1}", classificator.FamilyName, classificator.TypeName), Output.OutputMessageType.Code, debug);
            }
            output.PrintDebug(string.Format("Заполнение классификатора по {0} ↑", storage.instanceOrType == 1 ? "экземпляру" : "типу"), Output.OutputMessageType.Header, debug);

            foreach (Element elem in constrs)
            {
                if (elem == null)
                {
                    output.PrintInfo("Не удалось получить элементы для заполнения классификатора! " +
                        "Проверь выборку (можно выбирать только моделируемые элементы), либо запусти на весь проект!", Output.OutputMessageType.Error);
                    return false;
                }
                
                string familyName = getElemFamilyName(elem);

                output.PrintDebug(string.Format("{0} : {1} : {2}", elem.Name, familyName, elem.Id.IntegerValue), Output.OutputMessageType.Regular, debug);
                foreach (Classificator classificator in storage.classificator)
                {
                    bool categoryCatch = false;
                    try
                    {
                        categoryCatch = Category.GetCategory(doc, classificator.BuiltInName).Id.IntegerValue == elem.Category.Id.IntegerValue;
                    }
                    catch (Exception)
                    {
                        output.PrintDebug(string.Format("Не удалось определить категорию из файла классификатора: {0}. Возможно, она введена неверно.", classificator.BuiltInName), Output.OutputMessageType.Error, debug);
                    }
                    bool familyNameCatch = nameChecker(classificator.FamilyName, familyName);
                    bool typeNameCatch = nameChecker(classificator.TypeName, elem.Name);
                    bool familyNameNotExist = classificator.FamilyName.Length == 0;
                    bool typeNameNotExist = classificator.TypeName.Length == 0;

                    if (categoryCatch && familyNameCatch && typeNameCatch && !familyNameNotExist && !typeNameNotExist)
                    {
                        setClassificator(classificator, storage, elem);
                        break;
                    }
                    if (categoryCatch && familyNameCatch && !familyNameNotExist && typeNameNotExist)
                    {
                        setClassificator(classificator, storage, elem);
                        break;
                    }
                    if (categoryCatch && typeNameCatch && familyNameNotExist && !typeNameNotExist)
                    {
                        setClassificator(classificator, storage, elem);
                        break;
                    }
                    if (categoryCatch && familyNameNotExist && typeNameNotExist)
                    {
                        setClassificator(classificator, storage, elem);
                        break;
                    }
                }
            }
            return true;
        }

        public static string getElemFamilyName(Element elem)
        {
            string familyName = null;
            if (elem is Room)
            {
                familyName = elem.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString();
            }
            else
            {
                familyName = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                familyName = familyName == null || familyName.Length == 0 ? (elem as ElementType).FamilyName : familyName;
            }

            if (familyName == null)
                throw new Exception("Не удалось взять имя семейства! Обратись к разработчику");
            
            return familyName;
        }
    }
}
