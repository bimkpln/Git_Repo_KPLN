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
using static KPLN_Classificator.ApplicationConfig;
using Autodesk.Revit.UI.Selection;
using KPLN_Classificator.Data;
using KPLN_Classificator.Forms;

namespace KPLN_Classificator
{
    class CommandStartClassificator : MyExecutableCommand
    {
        private ClassificatorForm form { get; set; }

        public CommandStartClassificator(ClassificatorForm form)
        {
            this.form = form;
        }
        public Result Execute(UIApplication app)
        {
            Document doc = app.ActiveUIDocument.Document;
            Selection sel = app.ActiveUIDocument.Selection;
            View activeView = doc.ActiveView;

            bool debugMode = form.debugMode;
            bool colourMode = form.colourMode;

            InfosStorage storage = form.storage;
            if (storage == null)
            {
                output.PrintInfo("Не удалось обработать конфигурационный файл!", Output.OutputMessageType.Error);
                return Result.Cancelled;
            }

            List<BuiltInCategory> constrCats = form.checkedCats;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Заполнение параметров классификатора");

                if (form.instanceOrType == 2)
                {
                    List<Element> constrsTypes;

                    // Обработка всех элементов модели
                    if (sel.GetElementIds().Count == 0)
                    {
                        constrsTypes = new FilteredElementCollector(doc).WhereElementIsElementType()
                            .WherePasses(new ElementMulticategoryFilter(constrCats))
                            .ToElements()
                            .Where(i => i.get_Parameter(BuiltInParameter.UNIFORMAT_CODE) != null)
                            .ToList();
                    }
                    // Обработка только выбранных элементов
                    else
                    {
                        constrsTypes = sel.GetElementIds()
                            .Where(id => doc.GetElement(id).Category.CategoryType == CategoryType.Model)
                            .Select(elem => doc.GetElement(elem).GetTypeId())
                            .Select(type => doc.GetElement(type))
                            .ToHashSet()
                            .ToList();
                    }

                    ParamUtils utilsForType = new ParamUtils(debugMode);
                    if (!utilsForType.startClassification(constrsTypes, storage, doc))
                    {
                        return Result.Cancelled;
                    }
                    output.PrintInfo(string.Format("Успешно обработано элементов: {0}. Обработано с ошибками: {1}.",
                        utilsForType.fullSuccessElems.Count, utilsForType.notFullSuccessElems.Count), Output.OutputMessageType.Success);
                }

                else if (form.instanceOrType == 1)
                {
                    List<Element> constrsInstances;

                    // Обработка всех элементов модели
                    if (sel.GetElementIds().Count == 0)
                    {
                        constrsInstances = new FilteredElementCollector(doc)
                            .WhereElementIsNotElementType()
                            .WherePasses(new ElementMulticategoryFilter(constrCats))
                            .ToList();
                    }
                    // Обработка только выбранных элементов
                    else
                    {
                        constrsInstances = sel.GetElementIds()
                            .Where(id => doc.GetElement(id).Category.CategoryType == CategoryType.Model)
                            .Select(elem => doc.GetElement(elem))
                            .ToList();
                        List<Element> constrsInstancesFromGroup = new List<Element>();
                        foreach (Element item in constrsInstances)
                        {
                            if (item is Group)
                            {
                                Group group = (Group)item;
                                if (group == null) continue;
                                constrsInstancesFromGroup.AddRange(group.GetMemberIds().Select(elem => doc.GetElement(elem)).ToList());
                            }
                        }
                        constrsInstances.AddRange(constrsInstancesFromGroup);
                    }

                    ParamUtils utilsForInstanse = new ParamUtils(debugMode);
                    if (!utilsForInstanse.startClassification(constrsInstances, storage, doc))
                    {
                        return Result.Cancelled;
                    }

                    if (colourMode && activeView.ViewType == ViewType.ThreeD)
                    {
                        ViewUtils viewUtils = new ViewUtils(doc);

                        OverrideGraphicSettings overrideGraphic = ViewUtils.getStandartGraphicSettings(doc);
                        List<Element> elems = new FilteredElementCollector(doc, activeView.Id)
                             .WhereElementIsNotElementType()
                             .ToElements()
                             .ToList();

                        foreach (Element elem in elems)
                        {
                            if (elem is Group) continue;
                            try
                            {
                                activeView.SetElementOverrides(elem.Id, overrideGraphic);
                            }
                            catch { }
                        }
                        foreach (Element elem in utilsForInstanse.fullSuccessElems)
                        {
                            if (elem is Group) continue;
                            try
                            {
                                activeView.SetElementOverrides(elem.Id, viewUtils.fullSuccessSet);
                            }
                            catch { }
                        }
                        foreach (Element elem in utilsForInstanse.notFullSuccessElems)
                        {
                            if (elem is Group) continue;
                            try
                            {
                                activeView.SetElementOverrides(elem.Id, viewUtils.notFullSuccessSet);
                            }
                            catch { }
                        }
                    }
                    output.PrintInfo(string.Format("Успешно обработано элементов: {0}. Обработано с ошибками: {1}.",
                        utilsForInstanse.fullSuccessElems.Count, utilsForInstanse.notFullSuccessElems.Count), Output.OutputMessageType.Success);
                }
                else
                {
                    output.PrintInfo("Выбрана некорректная операция! Проверьте конфигурационный файл.", Output.OutputMessageType.Error);
                    return Result.Cancelled;
                }

                try
                {
                    LastRunInfo.getInstanceOrCreateNew(doc).save(form.utils.xmlFilePath != null ? form.utils.xmlFilePath : "Файл конфигурации не был сохранён при последнем запуске!");
                }
                catch (Exception e)
                {
                    output.PrintErr(e, "Произошла ошибка в процессе сохранения информации о запуске плагина!");
                }

                t.Commit();
            }

            string fileName = string.Format("C:\\TEMP\\log_{0}.txt", DateTime.Now.ToString().Replace(".", "").Replace(":", ""));
            using (StreamWriter streamWriter = new StreamWriter(fileName))
            {
                try
                {
                    streamWriter.WriteLine(string.Format("Классификация файла: {0}", doc.Title));
                    streamWriter.WriteLine(output.getLog());
                    output.PrintInfo(string.Format("Отчёт о работе плагина сохранён в файле: {0}", fileName), Output.OutputMessageType.Regular);
                }
                catch (Exception e)
                {
                    output.PrintErr(e, "Не удалось выполнить запись ЛОГ файла!");
                }
            }
            output.clearLog();
            return Result.Succeeded;
        }
    }
}
