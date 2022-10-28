using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.Tools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandCheckElementWorksets : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                HashSet<string> links = new HashSet<string>();
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                Document doc = commandData.Application.ActiveUIDocument.Document;
                if (doc.IsWorkshared)
                {
                    List<Workset> worksets = new List<Workset>();
                    foreach (Workset w in new FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))
                    {
                        if (!w.IsOpen)
                        {
                            Print("Необходимо открыть все рабочие наборы!", KPLN_Loader.Preferences.MessageType.Error);
                            return Result.Cancelled;
                        }
                        worksets.Add(w);
                    }
                    foreach (Element element in new FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())
                    {
                        if (element.Category == null) { continue; }
                        
                        try
                        {
                            if (element.GetType() == typeof(RevitLinkInstance) || element.GetType() == typeof(ImportInstance))
                            {
                                continue;
                            }
                            
                            if ((element.Category.CategoryType == CategoryType.Annotation) & (element.GetType() == typeof(Grid) | element.GetType() == typeof(Level)))
                            {
                                string wsName = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString();
                                if (!wsName.ToLower().Contains("оси и уровни") & !wsName.ToLower().Contains("общие уровни и сетки"))
                                {
                                    WPFDisplayItem item = GetItemByElement(element, element.Name, "Ошибка рабочего набора", string.Format("Ось или уровень с ID {0} находится не в специальном рабочем наборе", element.Id), Status.Error, null);
                                    item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Подсказка: ", Description = "Имя рабочего набора для осей и уровней - <..._Оси и уровни>" });
                                    outputCollection.Add(item);
                                }    
                            }
                            
                            // Есть внутренняя ошибка Revit, когда появляются компоненты легенды, которые нигде не размещены, и у них редактируемый рабочий набор. Вручную такой элемент - создать НЕВОЗМОЖНО
                            if (element.Category.CategoryType == CategoryType.Model && element.Category.Id.IntegerValue != -2000576)
                            {
                                foreach (Workset w in worksets)
                                {
                                    if (element.WorksetId.IntegerValue == w.Id.IntegerValue)
                                    {
                                        // Проверка замонитренных моделируемых элементов
                                        if (element.GetMonitoredLinkElementIds().Count() > 0)
                                        {
                                            if (!w.Name.StartsWith("02"))
                                            {
                                                UpdateOutputCollection(
                                                    element,
                                                    w,
                                                    "Элементы с мониторингом (т.е. скопированные из других файлов) должны находится в рабочих наборах с приставкой '02'",
                                                    ref outputCollection);

                                                continue;
                                            }
                                        }
                                        
                                        // Проверка остальных моделируемых элементов на рабочий набор связей 
                                        else if (w.Name.StartsWith("00")
                                            | w.Name.StartsWith("#")
                                            && !w.Name.Contains("DWG"))
                                        {
                                            UpdateOutputCollection(
                                                element,
                                                w,
                                                "В рабочих наборах связей должны быть только экземпляры rvt-связей",
                                                ref outputCollection);

                                            continue;
                                        }
                                        
                                        // Проверка остальных моделируемых элементов на рабочий набор для сеток
                                        else if (w.Name.ToLower().Contains("оси и уровни")
                                            | w.Name.ToLower().Contains("общие уровни и сетки"))
                                        {
                                            UpdateOutputCollection(
                                                element,
                                                w,
                                                "В рабочих наборах для осей и уровней должны быть только оси или уровни",
                                                ref outputCollection);

                                            continue;
                                        }
                                        
                                        // Проверка остальных моделируемых элементов на рабочий набор для связей
                                        else if (w.Name.StartsWith("02"))
                                        {
                                            UpdateOutputCollection(
                                                element, 
                                                w,
                                                "Только элементы с мониторингом должны быть в рабочих наборах с приставкой '02'", 
                                                ref outputCollection);

                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        { }
                    }
                }
                else
                {
                    WPFDisplayItem item = GetItemByElement(doc.Title, "Ошибка проекта", "Файл не настроен для совместной работы", Status.Error);
                    outputCollection.Add(item);
                }
                ObservableCollection<WPFDisplayItem> wpfCategories = new ObservableCollection<WPFDisplayItem>();
                wpfCategories.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Name = "<Все>" });
                List<WPFDisplayItem> sortedOutputCollection = outputCollection.OrderBy(o => o.Header).ToList();
                ObservableCollection<WPFDisplayItem> wpfElements = new ObservableCollection<WPFDisplayItem>();
                int counter = 1;
                foreach (WPFDisplayItem e in sortedOutputCollection)
                {
                    e.Header = string.Format("{0}# {1}", (counter++).ToString(), e.Header);
                    wpfElements.Add(e);
                }
                if (wpfElements.Count != 0)
                {
                    ElementsOutputExtended form = new ElementsOutputExtended(wpfElements, wpfCategories);
                    form.Show();
                }
                else
                {
                    Print("[Рабочие наборы] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }
        
        private void UpdateOutputCollection (Element element, Workset w, string description, ref ObservableCollection<WPFDisplayItem> outputCollection)
        {
            WPFDisplayItem item = GetItemByElement(element, element.Name, "Ошибка рабочего набора", string.Format("Элемент модели с ID {0} находится не в своем рабочем наборе «{1}»", element.Id, w.Name), Status.Error, null);

            item.Collection.Add(
                new WPFDisplayItem(-1, StatusExtended.Critical)
                {
                    Header = "Подсказка: ",
                    Description = description
                });

            outputCollection.Add(item);
        }

        private WPFDisplayItem GetItemByElement(string name, string header, string description, Status status)
        {
            StatusExtended exstatus;
            switch (status)
            {
                case Status.Error:
                    exstatus = StatusExtended.Critical;
                    break;
                default:
                    exstatus = StatusExtended.Warning;
                    break;
            }
            WPFDisplayItem item = new WPFDisplayItem(-1, exstatus);
            try
            {
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", "Документ");
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(-1, exstatus) { Header = "Описание: ", Description = description });
                HashSet<string> values = new HashSet<string>();
            }
            catch (Exception e)
            {
                try
                {
                    PrintError(e.InnerException);
                }
                catch (Exception) { }
                PrintError(e);
            }
            return item;
        }

        private WPFDisplayItem GetItemByElement(Element element, string name, string header, string description, Status status, BoundingBoxXYZ box)
        {
            StatusExtended exstatus;
            switch (status)
            {
                case Status.Error:
                    exstatus = StatusExtended.Critical;
                    break;
                default:
                    exstatus = StatusExtended.Warning;
                    break;
            }
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus);
            try
            {
                item.SetZoomParams(element, box);
                item.Name = name;
                item.Header = header;
                item.Description = description;
                item.Category = string.Format("<{0}>", element.Category.Name);
                item.Visibility = System.Windows.Visibility.Visible;
                item.IsEnabled = true;
                item.Collection = new ObservableCollection<WPFDisplayItem>();
                item.Collection.Add(new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus) { Header = "Подсказка: ", Description = description });
                HashSet<string> values = new HashSet<string>();
            }
            catch (Exception e)
            {
                try
                {
                    PrintError(e.InnerException);
                }
                catch (Exception) { }
                PrintError(e);
            }
            return item;
        }
    }
}
