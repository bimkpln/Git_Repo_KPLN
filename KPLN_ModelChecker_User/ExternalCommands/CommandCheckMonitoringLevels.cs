using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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
    public class CommandCheckMonitoringLevels : IExternalCommand
    {
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                List<WPFElement> Elements = new List<WPFElement>();
                Document doc = commandData.Application.ActiveUIDocument.Document;
                HashSet<int> ids = new HashSet<int>();
                ObservableCollection<WPFDisplayItem> outputCollection = new ObservableCollection<WPFDisplayItem>();
                foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
                {
                    try
                    {
                        if (element.IsMonitoringLinkElement())
                        {
                            RevitLinkInstance link = null;
                            List<string> names = new List<string>();
                            foreach (ElementId i in element.GetMonitoredLinkElementIds())
                            {
                                ids.Add(i.IntegerValue);
                                link = commandData.Application.ActiveUIDocument.Document.GetElement(i) as RevitLinkInstance;
                                names.Add(link.Name);
                            }
                            if (link == null)
                            {
                                outputCollection.Add(GetItemByElement(element, element.Name, "Связь не найдена", string.Format("Элементу с ID {0} необходимо исправить мониторинг", element.Id), Status.Error, null));
                            }
                        }
                        else
                        {
                            outputCollection.Add(GetItemByElement(element, element.Name, "Отсутствует мониторинг", string.Format("Элементу с ID {0} необходимо задать мониторинг", element.Id), Status.Error, null));
                        }
                    }
                    catch (Exception)
                    { }
                }
                if (ids.Count > 1)
                {
                    WPFDisplayItem item = GetItemByElement(doc.Title, "Мониторинг настроен на основе нескольких файлов", "Необходимо настраивать мониторинг на основе одного файла (Разбивочный файл/Архитектурный файл)", Status.Error);
                    foreach (int i in ids)
                    {
                        try
                        {
                            RevitLinkInstance link = doc.GetElement(new ElementId(i)) as RevitLinkInstance;
                            item.Collection.Add(new WPFDisplayItem(link.Category.Id.IntegerValue, StatusExtended.Critical) { Header = "Связь с мониторингом:", Description = string.Format("«{0}» <{1}>", link.Name, link.Id.ToString()) });

                        }
                        catch (Exception)
                        {
                            item.Collection.Add(new WPFDisplayItem(-1, StatusExtended.Critical) { Header = "Связь с мониторингом:", Description = string.Format("Не найдена <{0}>", i.ToString()) });
                        }
                    }
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
                    Print("[Мониторинг] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }
    }
}
