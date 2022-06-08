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
                string folder = Path.GetDirectoryName(dllPath);
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

                    //Назначение рабочих наборов для элементов модели
                    string department = dto.Department;
                    string linkedFilesPrefix = dto.LinkedFilesPrefix;
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
                                    .Where(f => f.Symbol.FamilyName.ToLower().Contains(familyName.ToLower()))
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

                                    if (elemType.Name.ToLower().Contains(typeName.ToLower()))
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
                                        string data = elem.LookupParameter(p.ParameterName).AsString();
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
                        bool isUnique = WorksetTable.IsWorksetNameUnique(doc, linkWorksetName);
                        if (isUnique)
                        {
                            Workset.Create(doc, linkWorksetName);
                        }

                        Workset linkWorkset = new FilteredWorksetCollector(doc)
                            .OfKind(WorksetKind.UserWorkset)
                            .ToWorksets()
                            .Where(w => w.Name == linkWorksetName)
                            .First();

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
    }
}
