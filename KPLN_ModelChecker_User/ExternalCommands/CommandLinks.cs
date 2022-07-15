using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
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
    public class CommandLinks : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                Document doc = commandData.Application.ActiveUIDocument.Document;

                // Обрабатываю rvt-связи
                IList<Element> rvtLinks = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements();
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
                    foreach (RevitLinkInstance link in rvtLinks)
                    {
                        string[] separators = { ".rvt : " };
                        string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                        int lenNS = nameSubs.Length;
                        if (lenNS > 3)
                        {
                            continue;
                        }
                        foreach (Workset w in worksets)
                        {
                            if (link.WorksetId.IntegerValue == w.Id.IntegerValue)
                            {
                                if (!w.Name.StartsWith("00") & !w.Name.StartsWith("#"))
                                {
                                    WPFDisplayItem item = GetItemByElement(link, link.Name, "Ошибка рабочего набора", "Связь находится в некорректном рабочем наборе", Status.Error, null);
                                    item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Подсказка: ", Description = "Рабочие наборы связей должны соответствовать формату 00_<Раздел>_<Подраздел>_Модель, напр.: «00_КР_К2_Модель»" });
                                    outputCollection.Add(item);
                                }
                            }
                        }
                    }
                }
                else
                {
                    WPFDisplayItem item = GetItemByElement(doc.Title, "Ошибка проекта", "Файл не настроен для совместной работы", Status.Error);
                    outputCollection.Add(item);
                }
                
                foreach (RevitLinkInstance link in rvtLinks)
                {
                    string[] separators = { ".rvt : " };
                    string[] nameSubs = link.Name.Split(separators, StringSplitOptions.None);
                    int lenNS = nameSubs.Length;
                    if (lenNS > 2) 
                    { 
                        continue; 
                    }
                    try
                    {
                        Document linkDocument = link.GetLinkDocument();
                        string name = linkDocument.PathName;
                        string currentPosition = link.Name.Split(new string[] { "позиция " }, StringSplitOptions.RemoveEmptyEntries).Last();
                        /*
                        Это была проверка на дубликаты экземпляров связей. Ревит не позволяет нескольким экземплярам связей иметь одну и ту же площадку. Т.е. такая связь попадет в ошибку с отсутствием площадки
                        if (!links.Contains(name))
                        {
                            links.Add(name);
                        }
                        else
                        {
                            WPFDisplayItem item = GetItemByElement(link, link.Name, "Дублирование", "Несколько экземпляров одной и той же подгруженной связи", Status.Error, null);
                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                            outputCollection.Add(item);
                        }
                        */
                        if (currentPosition == "<Не общедоступное>")
                        {
                            WPFDisplayItem item = GetItemByElement(link, link.Name, "Площадка отсутствует", "Для связи необходимо назначить общую площадку", Status.Error, null);
                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                            if (link.Pinned)
                            {
                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✔ - Закреплена" });
                            }
                            else
                            {
                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✘ - Не закреплена" });
                            }
                            outputCollection.Add(item);
                        }
                        else
                        {
                            bool detected = false;
                            foreach (ProjectLocation i in doc.ProjectLocations)
                            {
                                if (i.Name == currentPosition)
                                {
                                    detected = true;
                                    if (currentPosition == "Встроенный")
                                    {
                                        WPFDisplayItem item = GetItemByElement(link, link.Name, "Площадка настроена неправильно", "Запрещено использовать системное имя для площадки", Status.Error, null);
                                        item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                                        if (link.Pinned)
                                        {
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✔ - Закреплена" });
                                        }
                                        else
                                        {
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✘ - Не закреплена" });
                                        }
                                        outputCollection.Add(item);
                                    }
                                    else
                                    {
                                        if (!link.Pinned)
                                        {
                                            WPFDisplayItem item = GetItemByElement(link, link.Name, "Площадка настроена", "Связь не закреплена", Status.Error, null);
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✘ - Не закреплена" });
                                            outputCollection.Add(item);
                                        }
                                    }
                                }
                            }
                            if (!detected)
                            {
                                foreach (ProjectLocation i in linkDocument.ProjectLocations)
                                {
                                    if (i.Name == currentPosition)
                                    {
                                        detected = true;
                                        if (currentPosition == "Встроенный")
                                        {
                                            WPFDisplayItem item = GetItemByElement(link, link.Name, "Площадка настроена неправильно", "Запрещено использовать системное имя для площадки", Status.Error, null);
                                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                                            if (link.Pinned)
                                            {
                                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✔ - Закреплена" });
                                            }
                                            else
                                            {
                                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✘ - Не закреплена" });
                                            }
                                            outputCollection.Add(item);
                                        }
                                        else
                                        {
                                            if (!link.Pinned)
                                            {
                                                WPFDisplayItem item = GetItemByElement(link, link.Name, "Площадка настроена", "Связь не закреплена", Status.Error, null);
                                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = currentPosition });
                                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "✘ - Не закреплена" });
                                                outputCollection.Add(item);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        try
                        {
                            WPFDisplayItem item = GetItemByElement(link, link.Name, "Связь не подгружена", "Ошибка проверки координат. Необходимо подгрузить связь в проект", Status.Error, null);
                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Текущая позиция: ", Description = "-" });
                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Закрепление: ", Description = "-" });
                            outputCollection.Add(item);
                        }
                        catch (Exception) { }
                    }
                }

                // Заполняю пользовательское окно результатов
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
                    Print("[Мониторинг связей] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
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
            WPFDisplayItem item = new WPFDisplayItem(-1, exstatus, "✔");
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
            WPFDisplayItem item = new WPFDisplayItem(element.Category.Id.IntegerValue, exstatus, "✔");
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
