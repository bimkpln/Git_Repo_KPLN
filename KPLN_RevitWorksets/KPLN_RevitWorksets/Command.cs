#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB; 
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Xml;
using static KPLN_Loader.Output.Output;

namespace KPLN_RevitWorksets
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {

                Document doc = commandData.Application.ActiveUIDocument.Document;

                if (!doc.IsWorkshared)
                {
                    message = "Файл не является файлом совместной работы";
                    return Result.Failed; ;
                }

                //считываю список рабочих наборов
                string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string folder = System.IO.Path.GetDirectoryName(dllPath);

                System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
                dialog.InitialDirectory = folder;
                dialog.Multiselect = false;
                dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return Result.Cancelled;
                string xmlFilePath = dialog.FileName;

                InfosStorage storage = new InfosStorage();
                System.Xml.Serialization.XmlSerializer serializer =
                    new System.Xml.Serialization.XmlSerializer(typeof(InfosStorage));
                using (StreamReader r = new StreamReader(xmlFilePath))
                {
                    storage = (InfosStorage)serializer.Deserialize(r);
                }
                if (storage.LinkedFilesPrefix == null) storage.LinkedFilesPrefix = "#";

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Создание рабочих наборов");

                    //назначаю рабочие наборы в случае, если указана категория
                    foreach (WorksetByCategory wb in storage.worksetsByCategory)
                    {
                        Workset wset = wb.GetWorkset(doc);
                        List<BuiltInCategory> cats = wb.revitCategories;
                        if (cats == null) continue;
                        if (cats.Count == 0) continue;

                        foreach (BuiltInCategory bic in cats)
                        {
                            List<Element> elems = new FilteredElementCollector(doc)
                                .OfCategory(bic)
                                .WhereElementIsNotElementType()
                                .ToElements()
                                .ToList();

                            foreach (Element elem in elems)
                            {
                                WorksetBy.SetWorkset(elem, wset);
                            }
                        }
                    }

                    //назначаю рабочие наборы по именам семейств
                    List<FamilyInstance> famIns = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .ToList();
                    foreach (WorksetByFamily wb in storage.worksetsByFamily)
                    {
                        Workset wset = wb.GetWorkset(doc);

                        List<string> families = wb.FamilyNames;
                        if (families == null) continue;
                        if (families.Count == 0) continue;



                        foreach (string familyName in families)
                        {
                            List<FamilyInstance> curFamIns = famIns
                                .Where(f => f.Symbol.FamilyName.ToLower().Contains(familyName.ToLower()))
                                .ToList();

                            foreach (FamilyInstance fi in curFamIns)
                            {
                                WorksetBy.SetWorkset(fi, wset);
                            }
                        }
                    }

                    //назначаю рабочие наборы по именам типов
                    List<Element> allElems = new FilteredElementCollector(doc)
                                .WhereElementIsNotElementType()
                                .Cast<Element>()
                                .ToList();

                    foreach (WorksetByType wb in storage.worksetsByType)
                    {
                        Workset wset = wb.GetWorkset(doc);
                        List<string> typeNames = wb.TypeNames;
                        if (typeNames == null) continue;
                        if (typeNames.Count == 0) continue;

                        foreach (string typeName in typeNames)
                        {
                            foreach (Element elem in allElems)
                            {
                                ElementId typeId = elem.GetTypeId();
                                if (typeId == null || typeId == ElementId.InvalidElementId) continue;
                                ElementType elemType = doc.GetElement(typeId) as ElementType;
                                if (elemType == null) continue;

                                if (elemType.Name.ToLower().Contains(typeName.ToLower()))
                                {
                                    WorksetBy.SetWorkset(elem, wset);
                                }
                            }
                        }
                    }

                    //назначаю рабочие наборы по значению параметров
                    foreach (WorksetByParameter wb in storage.worksetsByParameter)
                    {
                        Workset wset = wb.GetWorkset(doc);
                        string paramName = wb.ParameterName;
                        string paramValue = wb.ParameterValue;

                        List<Element> elemsFilterByParam = allElems
                            .Where(e => e.LookupParameter(paramName).AsString().Equals(paramValue)).ToList();

                        foreach (Element elem in elemsFilterByParam)
                        {
                            WorksetBy.SetWorkset(elem, wset);
                        }
                    }


                    //назначаю рабочие наборы для связанных файлов
                    List<RevitLinkInstance> links = new FilteredElementCollector(doc)
                        .OfClass(typeof(RevitLinkInstance))
                        .Cast<RevitLinkInstance>()
                        .ToList();

                    foreach (RevitLinkInstance rli in links)
                    {
                        RevitLinkType linkFileType = doc.GetElement(rli.GetTypeId()) as RevitLinkType;
                        if (linkFileType == null) continue;
                        if (linkFileType.IsNestedLink) continue;

                        string linkWorksetName1 = rli.Name.Split(':')[0];
                        string linkWorksetName2 = linkWorksetName1.Substring(0, linkWorksetName1.Length - 5);
                        string linkWorksetName = storage.LinkedFilesPrefix + linkWorksetName2;
                        bool checkExists = WorksetTable.IsWorksetNameUnique(doc, linkWorksetName);
                        if (!checkExists) continue;

                        Workset.Create(doc, linkWorksetName);

                        Workset linkWorkset = new FilteredWorksetCollector(doc)
                            .OfKind(WorksetKind.UserWorkset)
                            .ToWorksets()
                            .Where(w => w.Name == linkWorksetName)
                            .First();

                        WorksetBy.SetWorkset(rli, linkWorkset);
                        WorksetBy.SetWorkset(linkFileType, linkWorkset);
                    }

                    t.Commit();
                }

                List<string> emptyWorksetsNames = WorksetTool.GetEmptyWorksets(doc);
                if (emptyWorksetsNames.Count > 0)
                {
                    string msg = "Обнаружены пустые рабочие наборы! Их можно удалить вручную:\n";
                    foreach (string s in emptyWorksetsNames)
                    {
                        msg += s + "\n";
                    }
                    TaskDialog.Show("Отчёт", msg);
                }

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                PrintError(e, "Произошла ошибка во время запуска скрипта");
                return Result.Failed;
            }
        }
    }
}
