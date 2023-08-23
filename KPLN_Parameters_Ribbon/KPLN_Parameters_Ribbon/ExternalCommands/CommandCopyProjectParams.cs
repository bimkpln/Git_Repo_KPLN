using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;

namespace KPLN_Parameters_Ribbon.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    class CommandCopyProjectParams : IExternalCommand
    {
        private readonly List<string> _parametersName = new List<string>()
        {
            "SHT_Вид строительства",
            "SHT_Абсолютная отметка",
            "Дата утверждения проекта",
            "Статус проекта",
            "Заказчик",
            "Адрес проекта",
            "Наименование проекта",
            "Номер проекта",
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Document currentDoc = commandData.Application.ActiveUIDocument.Document;
                DocumentSet activeDocs = commandData.Application.Application.Documents;
                Document baseDoc = null;

                if (activeDocs.Size < 2)
                {
                    throw new Exception("Для копирования параметров необходимо открыть файл, содержащий сведения о проекте!");
                }

                foreach (Document doc in activeDocs)
                {
                    if (doc.Title.ToLower().Contains("сведения"))
                    {
                        baseDoc = doc;
                        break;
                    }
                }

                if (baseDoc == null)
                {
                    throw new Exception("Файл, содержащий сведения о проекте не открыт!");
                }

                ProjectInfo baseInfo = new FilteredElementCollector(baseDoc)
                    .OfClass(typeof(ProjectInfo))
                    .OfCategory(BuiltInCategory.OST_ProjectInformation)
                    .Cast<ProjectInfo>()
                    .ToList()[0];

                ProjectInfo currentInfo = new FilteredElementCollector(currentDoc)
                    .OfClass(typeof(ProjectInfo))
                    .OfCategory(BuiltInCategory.OST_ProjectInformation)
                    .Cast<ProjectInfo>()
                    .ToList()[0];

                Dictionary<string, Parameter> paramsDict = new Dictionary<string, Parameter>();
                foreach (string currentName in _parametersName)
                {
                    paramsDict.Add(currentName, baseInfo.LookupParameter(currentName));
                }

                int counter = 0;
                int falseCounter = 0;

                using (Transaction t = new Transaction(currentDoc))
                {
                    t.Start("Копирование параметров");

                    Print(string.Format("Копирование параметров из файла: \"{0}\" в файл: \"{1}\" ↑", baseDoc.Title, currentDoc.Title), MessageType.Header);

                    foreach (KeyValuePair<string, Parameter> kvp in paramsDict)
                    {
                        if (CopyerParams(kvp, currentInfo))
                        {
                            counter++;
                        }
                        else
                        {
                            falseCounter++;
                        }
                    }
                    Print(string.Format("Скопировано параметров: {0}, не скопировано: {1}", counter, falseCounter), MessageType.Success);
                    t.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта!");
                return Result.Failed;
            }
        }

        public bool CopyingGlobalParams(GlobalParameter param, Document currentDoc)
        {
            bool check = false;
            if (param == null)
            {
                return check;
            }
            string paramName = param.GetDefinition().Name;
            try
            {
                ParameterValue paramValue = param.GetValue();
                if (paramValue == null)
                {
                    Print("Не скопирован параметр: " + paramName + ", т.к. он пуст.", MessageType.Code);
                    return check;
                }
                ElementId targetGlobalParamId = GlobalParametersManager.FindByName(currentDoc, paramName);
                GlobalParameter targetGlobalParam = currentDoc.GetElement(targetGlobalParamId) as GlobalParameter;
                targetGlobalParam.SetValue(paramValue);
                Print(string.Format("Параметру: \"{0}\" присвоено значение: \"{1}\"", paramName, Math.Round((paramValue as DoubleParameterValue).Value * 57.2957795D)), MessageType.Code);
                check = true;
            }
            catch (Exception)
            {
                Print(string.Format("Не удалось присвоить значение параметру: \"{0}\". Возможно, в файле: \"{1}\" данный параметр отсутствует.", paramName, currentDoc.Title), MessageType.Error);
            }
            return check;
        }

        public bool CopyerParams(KeyValuePair<string, Parameter> kvp, Element currentInfo)
        {
            bool check = false;
            string paramName = kvp.Key;
            Parameter param = kvp.Value;
            if (param == null)
            {
                Print("Не скопирован параметр: " + paramName + ", т.к. он отсутсвует. Обратись к норм. контроллеру.", MessageType.Warning);
                return check;
            }
            if (!param.HasValue)
            {
                Print("Не скопирован параметр: " + paramName + ", т.к. он пуст.", MessageType.Code);
                return check;
            }
            Parameter currentParam = currentInfo.LookupParameter(paramName);
            try
            {
                if (param.StorageType == StorageType.String)
                {
                    currentParam.Set(param.AsString());
                    Print(string.Format("Параметру: \"{0}\" присвоено значение: \"{1}\"", paramName, param.AsString()), MessageType.Code);
                    check = true;
                }
                else if (param.StorageType == StorageType.Double)
                {
                    currentParam.Set(param.AsDouble());
                    Print(string.Format("Параметру: \"{0}\" присвоено значение: \"{1}\"", paramName, param.AsDouble()), MessageType.Code);
                    check = true;
                }
                else if (param.StorageType == StorageType.Integer)
                {
                    currentParam.Set(param.AsInteger());
                    Print(string.Format("Параметру: \"{0}\" присвоено значение: \"{1}\"", paramName, param.AsInteger()), MessageType.Code);
                    check = true;
                }
                else
                {
                    Print("Не удалось определить тип параметра: " + paramName, MessageType.Error);
                }
            }
            catch (Exception e)
            {
                PrintError(e, "Не удалось присвоить параметр: " + paramName);
            }
            return check;
        }
    }
}
