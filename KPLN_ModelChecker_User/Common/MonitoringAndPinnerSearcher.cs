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
    internal static class MonitoringAndPinnerSearcher
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

        public static void GetLinks(ExternalCommandData commandData, Document doc, BuiltInCategory bic, ref ObservableCollection<WPFDisplayItem> outputCollection)
        {
            foreach (Element element in new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements())
            {
                if (element.IsMonitoringLinkElement())
                {
                    RevitLinkInstance link = null;
                    foreach (ElementId i in element.GetMonitoredLinkElementIds())
                    {
                        link = commandData.Application.ActiveUIDocument.Document.GetElement(i) as RevitLinkInstance;
                        if (link == null)
                        {
                            outputCollection.Add(
                                GetItemByElement(
                                    element,
                                    element.Name,
                                    $"Связь не найдена: «{element.Name}»",
                                    $"Элементу с ID {element.Id} необходимо исправить мониторинг",
                                    Status.Error,
                                    null)
                                );
                        }
                        else if (!link.Name.ToLower().Contains("разб"))
                        {
                            outputCollection.Add(
                                GetItemByElement(
                                    element,
                                    element.Name,
                                    $"Мониторинг не из разбивочного файла: «{element.Name}»",
                                    $"Элементу с ID {element.Id} необходимо исправить мониторинг, сейчас он присвоен связи {link.Name}",
                                    Status.Error,
                                    null)
                                );
                        }
                    }
                }
                else
                {
                    outputCollection.Add(
                        GetItemByElement(
                            element,
                            element.Name,
                            $"Отсутствует мониторинг: «{element.Name}»",
                            $"Элементу с ID {element.Id} необходимо задать мониторинг",
                            Status.Error,
                            null)
                        );
                }
                if (!element.Pinned)
                {
                    outputCollection.Add(
                        GetItemByElement(
                            element,
                            element.Name,
                            $"Элемент не прикреплен: «{element.Name}»",
                            $"Элемент с ID {element.Id} необходимо прикрепить",
                            Status.Error,
                            null)
                        );
                }
            }
        }
    }
}
