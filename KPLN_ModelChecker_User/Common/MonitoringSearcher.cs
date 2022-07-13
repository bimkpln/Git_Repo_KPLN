using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Common
{
    internal static class MonitoringSearcher
    {
        private static WPFDisplayItem GetItemByElement(string name, string header, string description, Status status)
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
                item.Collection.Add(new WPFDisplayItem(-1, exstatus) { Header = "Подсказка: ", Description = description });
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

        private static WPFDisplayItem GetItemByElement(Element element, string name, string header, string description, Status status, BoundingBoxXYZ box)
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

        public static void GetLinks(ExternalCommandData commandData, Document doc, BuiltInCategory bic, ref ObservableCollection<WPFDisplayItem> outputCollection, ref HashSet<int> ids)
        {
            foreach (Element element in new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
            {
                if (element.IsMonitoringLinkElement())
                {
                    RevitLinkInstance link = null;
                    foreach (ElementId i in element.GetMonitoredLinkElementIds())
                    {
                        ids.Add(i.IntegerValue);
                        link = commandData.Application.ActiveUIDocument.Document.GetElement(i) as RevitLinkInstance;
                    }
                    
                    if (link == null)
                    {
                        outputCollection.Add(
                            GetItemByElement(
                                element,
                                element.Name,
                                string.Format("Связь не найдена: Ось «{0}»", element.Name),
                                string.Format("Элементу с ID {0} необходимо исправить мониторинг", element.Id),
                                Status.Error,
                                null)
                            );
                    }
                    
                    else if (ids.Count > 1)
                    {
                        WPFDisplayItem item = GetItemByElement(
                            doc.Title,
                            "Мониторинг настроен на основе нескольких файлов",
                            "Необходимо настраивать мониторинг на основе одного файла (Разбивочный файл/Архитектурный файл)",
                            Status.Error
                        );
                        
                        foreach (int i in ids)
                        {
                            try
                            {
                                link = doc.GetElement(new ElementId(i)) as RevitLinkInstance;
                                item.Collection.Add(new WPFDisplayItem(link.Category.Id.IntegerValue, StatusExtended.Critical) 
                                    { 
                                        Header = "Связь с мониторингом:",
                                        Description = string.Format("«{0}» <{1}>", link.Name, link.Id.ToString())
                                    }
                                );

                            }
                            catch (Exception)
                            {
                                item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) 
                                    { 
                                        Header = "Связь с мониторингом:",
                                        Description = string.Format("Не найдена <{0}>", i.ToString()) 
                                    }
                                );
                            }
                        }
                        outputCollection.Add(item);
                    }
                    
                    else if (!link.Name.Contains("Разб"))
                    {
                        outputCollection.Add(
                            GetItemByElement(
                                element,
                                element.Name,
                                string.Format("Мониторинг не из разбивочного файла: Ось «{0}»", element.Name),
                                string.Format("Элементу с ID {0} необходимо исправить мониторинг", element.Id),
                                Status.Error,
                                null)
                            );
                    }
                }
                else
                {
                    outputCollection.Add(
                        GetItemByElement(
                            element,
                            element.Name,
                            string.Format("Отсутствует мониторинг: Ось «{0}»", element.Name),
                            string.Format("Элементу с ID {0} необходимо задать мониторинг", element.Id),
                            Status.Error,
                            null)
                        );
                }
            }
        }
    }
}
