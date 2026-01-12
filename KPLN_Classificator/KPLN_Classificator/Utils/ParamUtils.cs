using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Classificator.ApplicationConfig;

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

        public static bool paramChecker(string parameterName, string parameterValue, Element elem)
        {
            if (string.IsNullOrEmpty(parameterName) || string.IsNullOrEmpty(parameterValue))
                return true;

            Parameter parameter = elem.LookupParameter(parameterName) ?? elem.Document.GetElement(elem.GetTypeId()).LookupParameter(parameterName);
            if (parameter == null)
            {
                CurrentOutput.PrintDebug(string.Format("В элементе: \"{0}\" с id: {1} не найден параметр: \"{2}\".", elem.Name, elem.Id, parameterName), Output.OutputMessageType.Warning, debug);
                return false;
            }

            string elemParamValue = null;
            switch (parameter.StorageType)
            {
                case StorageType.Double:
                    elemParamValue = GetDoubleValue(parameter).ToString();
                    break;
                case StorageType.Integer:
                    elemParamValue = GetIntegerValue(parameter).ToString();
                    break;
                case StorageType.String:
                    elemParamValue = GetStringValue(parameter);
                    break;
                default:
                    CurrentOutput.PrintDebug("Не удалось определить тип параметра: " + parameter, Output.OutputMessageType.Error, debug);
                    break;
            }

            return CheckUserStringInput(parameterValue, elemParamValue);
        }

        public static bool nameChecker(string nameClafi, string nameElem)
        {
            if (string.IsNullOrEmpty(nameClafi))
                return true;

            return CheckUserStringInput(nameClafi, nameElem);
        }

        private static bool CheckUserStringInput(string userInput, string dataToCheck)
        {
            string[] arrayClafiAnd = userInput.ToLower().Split(',');
            int index = arrayClafiAnd.Length;

            if (index == 1)
            {
                if (arrayClafiAnd.First().StartsWith("!"))
                {
                    return !dataToCheck.ToLower().Contains(arrayClafiAnd.First().Replace("!", "")) && !dataToCheck.Contains("!");
                }

                if (arrayClafiAnd.First().Contains('|'))
                {
                    bool check = false;
                    string[] arrayClafiOr = arrayClafiAnd.First().Split('|');
                    for (int j = 0; j < arrayClafiOr.Length; j++)
                    {
                        if (dataToCheck.ToLower().Contains(arrayClafiOr[j]))
                        {
                            check = true;
                            break;
                        }
                    }
                    return check;
                }

                return dataToCheck.ToLower().Contains(arrayClafiAnd.First());
            }
            else if (index > 1)
            {
                for (int i = 0; i < index; i++)
                {
                    if (arrayClafiAnd[i].StartsWith("!"))
                    {
                        if (!dataToCheck.ToLower().Contains(arrayClafiAnd[i].Replace("!", "")) && !dataToCheck.Contains("!"))
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
                            if (dataToCheck.ToLower().Contains(arrayClafiOr[j]))
                            {
                                check = true;
                                break;
                            }
                        }
                        if (check) continue;
                        else return false;
                    }
                    else if (dataToCheck.ToLower().Contains(arrayClafiAnd[i])) continue;
                    else return false;
                }
                return true;
            }

            return false;
        }

        private static bool setParam(Element elem, string targetParamName, string value, out string valueForAssigned)
        {
            valueForAssigned = null;
            bool rsl = false;
            Parameter targetParam = elem.LookupParameter(targetParamName);
            if (targetParam == null)
            {
                CurrentOutput.PrintDebug(string.Format("В элементе: \"{0}\" с id: {1} не найден параметр: \"{2}\".", elem.Name, elem.Id, targetParamName), Output.OutputMessageType.Warning, debug);
                return rsl;
            }

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
                                    valueOfParam = GetValueStringOfParam_ByTargetParam(elem, foundParamNameForGettingValue.Split('*')[0], targetParam);
                                    if (valueOfParam == null) return rsl;
                                    try
                                    {
                                        valueOfParam = multipleSourseParamOnMultiplier(valueOfParam, foundParamNameForGettingValue);
                                    }
                                    catch (Exception)
                                    {
                                        CurrentOutput.PrintDebug(string.Format("Значение параметра: \"{0}\" в классификаторе содержит операцию умножения (*), которое не было выполнено. Проверьте корректность заполнения конфигурационного файла. Значение не вписано в параметр: \"{1}\".",
                                            foundParamName, targetParamName), Output.OutputMessageType.Warning, debug);
                                        return rsl;
                                    }
                                }
                                else
                                {
                                    valueOfParam = GetValueStringOfParam_ByTargetParam(elem, foundParamNameForGettingValue, targetParam);
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
                        CurrentOutput.PrintDebug(
                            $"Не заполнено значение параметра: \"{item}\" у элемента: {elem.Name} с id: {elem.Id}. Значение не вписано в параметр: \"{targetParamName}\".",
                            Output.OutputMessageType.Warning, debug);
                        return rsl;
                    }
                    newValue = newValue.Replace(item, itemValue);
                }
            }

            try
            {
                switch (targetParam.StorageType)
                {
                    case StorageType.Double:
                        targetParam.Set(double.Parse(newValue));
                        break;
                    case StorageType.Integer:
                        targetParam.Set(int.Parse(newValue));
                        break;
                    case StorageType.String:
                        targetParam.Set(newValue);
                        break;
                }
                valueForAssigned = newValue;
                rsl = true;
            }
            catch (Exception)
            {
                CurrentOutput.PrintDebug(string.Format("Не удалось присвоить значение \"{0}\" параметру: \"{1}\" с типом данных: {2}. Элемент: {3} с id: {4}",
                    newValue, targetParamName, targetParam.StorageType.ToString(), elem.Name, elem.Id), Output.OutputMessageType.Warning, debug);
            }
            return rsl;
        }

        private void setClassificator(Classificator classificator, InfosStorage storage, Element elem)
        {
            bool paramChecker;
            List<string> assignedValues = new List<string>();

            if (classificator.paramsValues.Count > storage.instanseParams.Count)
            {
                CurrentOutput.PrintDebug(string.Format("Значение параметра: \"{0}\" в элементе: \"{1}\" за пределами диапазона возможных значений. Присвоение данного параметра не будет выполнено."
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
            CurrentOutput.PrintDebug(string.Format("Были присвоены значения: {0}", string.Join("; ", assignedValues)), Output.OutputMessageType.System_OK, debug);
        }

        /// <summary>
        /// Взять значение текущего парамтера у элемента в формате string. Значение берем опираясь на параметр, в который этизначения будем записывать
        /// </summary>
        /// <param name="elem">Элемент модели</param>
        /// <param name="sourceParamName">Параметр, ИЗ которго забираем </param>
        /// <param name="targetParam">Параметр, В который вносим значения </param>
        /// <returns></returns>
        public static string GetValueStringOfParam_ByTargetParam(Element elem, string sourceParamName, Parameter targetParam)
        {
            string paramValue = null;

            Parameter sourceParam;
            if (elem is ElementType elemType)
                sourceParam = elem.LookupParameter(sourceParamName) ?? elemType.LookupParameter(sourceParamName);
            else
                sourceParam = elem.LookupParameter(sourceParamName) ?? elem.Document.GetElement(elem.GetTypeId()).LookupParameter(sourceParamName);
            
            if (sourceParam == null)
            {
                CurrentOutput.PrintDebug(string.Format("В элементе: \"{0}\" c id: {1} не найден параметр: \"{2}\"", elem.Name, elem.Id, sourceParamName), Output.OutputMessageType.Warning, debug);
                return paramValue;
            }

            switch (targetParam.StorageType)
            {
                case StorageType.Double:
                    paramValue = GetDoubleValue(sourceParam).ToString();
                    break;
                case StorageType.Integer:
                    paramValue = GetIntegerValue(sourceParam).ToString();
                    break;
                case StorageType.String:
                    paramValue = GetStringValue(sourceParam);
                    break;
                default:
                    CurrentOutput.PrintDebug("Не удалось определить тип параметра: " + sourceParamName, Output.OutputMessageType.Error, debug);
                    break;
            }

            return paramValue;
        }

        public bool startClassification(List<Element> constrs, InfosStorage storage, Document doc)
        {
            if (constrs == null || constrs.Count == 0)
            {
                CurrentOutput.PrintInfo("Не удалось получить элементы для заполнения классификатора! " +
                    "Проверь выборку (можно выбирать только моделируемые элементы), либо запусти на весь проект!", Output.OutputMessageType.Error);
                return false;
            }
            foreach (Classificator classificator in storage.classificator)
            {
                CurrentOutput.PrintDebug(string.Format("{0} - {1}", classificator.FamilyName, classificator.TypeName), Output.OutputMessageType.Code, debug);
            }
            CurrentOutput.PrintDebug(string.Format("Заполнение классификатора по {0} ↑", storage.instanceOrType == 1 ? "экземпляру" : "типу"), Output.OutputMessageType.Header, debug);

            foreach (Element elem in constrs)
            {
                if (elem == null)
                {
                    CurrentOutput.PrintInfo("Не удалось получить элементы для заполнения классификатора! " +
                        "Проверь выборку (можно выбирать только моделируемые элементы), либо запусти на весь проект!", Output.OutputMessageType.Error);
                    return false;
                }

                string familyName = getElemFamilyName(elem);

                CurrentOutput.PrintDebug(string.Format("{0} : {1} : {2}", elem.Name, familyName, elem.Id), Output.OutputMessageType.Regular, debug);
                foreach (Classificator classificator in storage.classificator)
                {
                    bool categoryCatch = false;
                    Category category = Category.GetCategory(doc, classificator.BuiltInName);
                    if (category != null)
                        categoryCatch = category.Id.Equals(elem.Category.Id);
                    else
                        CurrentOutput.PrintDebug(string.Format("Не удалось определить категорию из файла классификатора: {0}. Возможно, она введена неверно.", classificator.BuiltInName), Output.OutputMessageType.Error, debug);
                    if (!categoryCatch) continue;

                    bool familyNameCatch = nameChecker(classificator.FamilyName, familyName);
                    if (!familyNameCatch) continue;

                    bool typeNameCatch = nameChecker(classificator.TypeName, elem.Name);
                    if (!typeNameCatch) continue;

                    bool parameterValueCatch = paramChecker(classificator.ParameterName, classificator.ParameterValue, elem);
                    if (!parameterValueCatch) continue;

                    setClassificator(classificator, storage, elem);
                }
            }
            return true;
        }

        public static string getElemFamilyName(Element elem)
        {
            string familyName;
            if (elem is Room)
            {
                familyName = elem.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
            }
            else
            {
                familyName = elem.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM).AsValueString();
                if (elem is ElementType elemType)
                    familyName = familyName == null || familyName.Length == 0 ? elemType.FamilyName : familyName;
            }

            if (familyName == null)
                throw new Exception($"Не удалось взять имя семейства/помещения для элемента с id: {elem.Id}! Обратись к разработчику");

            return familyName;
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

        private static double? GetDoubleValue(Parameter p)
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

        private static int? GetIntegerValue(Parameter p)
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

        private static string GetStringValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.Double:
#if Revit2020 || Debug2020
                    return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), p.DisplayUnitType).ToString();
#endif
#if Revit2023 || Debug2023
                    return UnitUtils.ConvertFromInternalUnits(p.AsDouble(), p.GetUnitTypeId()).ToString();
#endif
                case StorageType.Integer:
                    return p.AsValueString();
                case StorageType.String:
                    return p.AsString();
                default:
                    return null;
            }
        }
    }
}
