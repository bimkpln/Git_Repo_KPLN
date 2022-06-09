using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Output;
using KPLN_ModelChecker_Debugger.ExternalCommands.Common;
using KPLN_ModelChecker_Debugger.ExternalCommands.Common.WorksetModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KPLN_ModelChecker_Debugger.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Worksetter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Получение объектов приложения и документа
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            try
            {
                if (!doc.IsWorkshared)
                {
                    message = "Файл не является файлом совместной работы";
                    return Result.Failed;
                }

                //Вывод пользовательского окна с xml-шаблонами
                string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string folder = $"X:\\BIM\\5_Scripts\\Git_Repo_KPLN\\KPLN_ModelChecker_Debugger\\KPLN_ModelChecker_Debugger\\Workset_Patterns";
                System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
                dialog.InitialDirectory = folder;
                dialog.Multiselect = false;
                dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Cancelled;
                }
                string xmlFilePath = dialog.FileName;

                //Десериализация пользовательского xml-файла
                WorksetDTO dto = new WorksetDTO();
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(typeof(WorksetDTO));
                using (StreamReader r = new StreamReader(xmlFilePath))
                {
                    dto = (WorksetDTO)serializer.Deserialize(r);
                }

                //Создание рабочих наборов
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("KPLN_Создание рабочих наборов");

                    // Общие поля класса WorksetDTO
                    string linkedFilesPrefix = dto.LinkedFilesPrefix;
                    bool useMonitoredElems = dto.UseMonitoredElements;
                    string monitoredElementsName = dto.MonitoredElementsName;

                    //Назначение рабочих наборов для элементов модели по правилам из поля класса WorksetDTO
                    foreach (WorksetByCurrentParameter param in dto.WorksetByCurrentParameterList)
                    {
                        Workset workset = param.GetWorkset(doc);

                        //Назначение рабочих наборов по категории
                        if (param.BuiltInCategories.Count != 0)
                        {
                            foreach (BuiltInCategory bic in param.BuiltInCategories)
                            {
                                List<Element> elems = new FilteredElementCollector(doc)
                                    .OfCategory(bic)
                                    .WhereElementIsNotElementType()
                                    .ToElements()
                                    .ToList();

                                foreach (Element elem in elems)
                                {
                                    WorksetByCurrentParameter.SetWorkset(elem, workset);
                                }
                            }
                        }

                        //Назначение рабочих наборов по имени семейства
                        if (param.FamilyNames.Count != 0)
                        {
                            List<FamilyInstance> famIns = new FilteredElementCollector(doc)
                                .WhereElementIsNotElementType()
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .ToList();
                            foreach (string familyName in param.FamilyNames)
                            {
                                List<FamilyInstance> elems = famIns
                                    .Where(f => f.Symbol.FamilyName.ToLower().StartsWith(familyName.ToLower()))
                                    .ToList();

                                foreach (Element elem in elems)
                                {
                                    WorksetByCurrentParameter.SetWorkset(elem, workset);
                                }
                            }
                        }

                        //Назначение рабочих наборов по имени типа
                        if (param.TypeNames.Count != 0)
                        {
                            List<Element> allModelElements = new FilteredElementCollector(doc)
                                .WhereElementIsNotElementType()
                                .Cast<Element>()
                                .ToList();
                            foreach (string typeName in param.TypeNames)
                            {
                                foreach (Element elem in allModelElements)
                                {
                                    ElementId typeId = elem.GetTypeId();
                                    if (typeId == null || typeId == ElementId.InvalidElementId) continue;
                                    
                                    ElementType elemType = doc.GetElement(typeId) as ElementType;
                                    if (elemType == null) continue;

                                    if (elemType.Name.ToLower().StartsWith(typeName.ToLower()))
                                    {
                                        WorksetByCurrentParameter.SetWorkset(elem, workset);
                                    }
                                }
                            }
                        }

                        //Назначение рабочих наборов по заполненному параметру
                        if (param.SelectedParameters.Count != 0)
                        {
                            List<Element> allModelElements = new FilteredElementCollector(doc)
                                .WhereElementIsNotElementType()
                                .Cast<Element>()
                                .ToList();

                            foreach (SelectedParameter p in param.SelectedParameters)
                            {
                                foreach (Element elem in allModelElements)
                                {
                                    try
                                    {
                                        string data;
                                        try
                                        {
                                            data = elem.LookupParameter(p.ParameterName).AsString();
                                        }
                                        catch (ArgumentNullException)
                                        {
                                            data = elem.LookupParameter(p.ParameterName).AsValueString();
                                        }
                                        if (data.Equals(p.ParameterValue))
                                        {
                                            WorksetByCurrentParameter.SetWorkset(elem, workset);
                                        }
                                    }
                                    catch (NullReferenceException) { }
                                }
                            }

                        }
                    }

                    //Назначение рабочих наборов для элементов с монитрингом (кроме осей и уровней)
                    if (useMonitoredElems)
                    {
                        Workset monitoredWorkset = CreateNewWorkset(doc, monitoredElementsName);
                        List<FamilyInstance> allModelElements = new FilteredElementCollector(doc)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .ToList();
                        
                        foreach (Element elem in allModelElements)
                        {
                            if ((elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Grids) 
                                && (elem.Category.Id.IntegerValue != (int)BuiltInCategory.OST_Levels)
                                && (elem.IsMonitoringLinkElement() || elem.IsMonitoringLocalElement()))
                            {
                                WorksetByCurrentParameter.SetWorkset(elem, monitoredWorkset);
                            }
                        }

                    }

                    //Назначение рабочих наборов для связанных файлов
                    FilteredElementCollector links = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance));
                    foreach (RevitLinkInstance linkInstance in links)
                    {
                        RevitLinkType linkFileType = doc.GetElement(linkInstance.GetTypeId()) as RevitLinkType;
                        if (linkFileType == null) continue;
                        if (linkFileType.IsNestedLink) continue;

                        string linkWorksetName1 = linkInstance.Name.Split(':')[0];
                        string linkWorksetName2 = linkWorksetName1.Substring(0, linkWorksetName1.Length - 5);
                        string linkWorksetName = linkedFilesPrefix + linkWorksetName2;
                        Workset linkWorkset = CreateNewWorkset(doc, linkWorksetName);

                        WorksetByCurrentParameter.SetWorkset(linkInstance, linkWorkset);
                        WorksetByCurrentParameter.SetWorkset(linkFileType, linkWorkset);
                    }

                   t.Commit();
                }

                List<string> emptyWorksetsNames = GetEmptyWorksets(doc);
                if (emptyWorksetsNames.Count > 0)
                {
                    string msg = "Обнаружены пустые рабочие наборы! Их следует удалить вручную:\n";
                    foreach (string s in emptyWorksetsNames)
                    {
                        msg += s + "\n";
                    }
                    TaskDialog.Show("Отчёт", msg);
                }

                return Result.Succeeded;
            }

            catch (Exception exc)
            {
                message = $"Произошла ошибка во время запуска скрипта - {exc}";
                return Result.Failed;
            }
        }

        /// <summary>
        /// Метод для поиска и вывода пользователю пустых рабочих наборов
        /// </summary>
        private List<string> GetEmptyWorksets(Document doc)
        {
            List<string> emptyWorksetsNames = new List<string>();
            if (!doc.IsWorkshared) return null;

            List<Workset> wids = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .ToList();

            foreach (Workset w in wids)
            {
                ElementWorksetFilter wfilter = new ElementWorksetFilter(w.Id);
                FilteredElementCollector col = new FilteredElementCollector(doc).WherePasses(wfilter);
                if (col.GetElementCount() == 0)
                {
                    emptyWorksetsNames.Add(w.Name);
                }
            }
            return emptyWorksetsNames;
        }

        /// <summary>
        /// Метод для создания рабочего набора
        /// </summary>
        private Workset CreateNewWorkset(Document doc, string name)
        {
            bool isUnique = WorksetTable.IsWorksetNameUnique(doc, name);
            if (isUnique)
            {
                Workset.Create(doc, name);
            }
            Workset linkWorkset = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets()
                .Where(w => w.Name == name)
                .First();
            return linkWorkset;
        }
    }
}
