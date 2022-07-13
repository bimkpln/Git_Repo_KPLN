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
                KPLN_ModelChecker_User.Common.MonitoringSearcher.GetLinks(commandData, doc, BuiltInCategory.OST_Levels, ref outputCollection, ref ids);
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
