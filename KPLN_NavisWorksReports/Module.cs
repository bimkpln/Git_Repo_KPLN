using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_NavisWorksReports.Tools;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using static KPLN_Loader.Output.Output;
using static KPLN_NavisWorksReports.ModuleData;

namespace KPLN_NavisWorksReports
{
    public class Module : IExternalModule
    {
        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
#if Revit2022
            MainWindowHandle = application.MainWindowHandle;
            HwndSource hwndSource = HwndSource.FromHwnd(MainWindowHandle);
            RevitWindow = hwndSource.RootVisual as Window;
#endif
#if Revit2020
            MainWindowHandle = application.MainWindowHandle;
            HwndSource hwndSource = HwndSource.FromHwnd(MainWindowHandle);
            RevitWindow = hwndSource.RootVisual as Window;
#endif
#if Revit2018
            try
            {
                MainWindowHandle = WindowHandleSearch.MainWindowHandle.Handle;
            }
            catch (Exception) { }
#endif
            string assembly = Assembly.GetExecutingAssembly().Location.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Split('.').First();
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "NavisWorks");
            AddPushButtonData(
                "NWC проверки",
                "Менеджер\nПроверок",
                "Утилита для работы с отчетами NavisWorks.",
                string.Format("{0}.{1}",
                    assembly,
                    "Commands.CommandShowManager"),
                panel,
                new Source.Source(Common.Collections.Icon.Default),
                true);
            return Result.Succeeded;
        }

        private void AddPushButtonData(string name, string text, string description, string className, RibbonPanel panel, Source.Source imageSource, bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = description;
            if (avclass)
            {
                button.AvailabilityClassName = "KPLN_NavisWorksReports.Availability.StaticAvailable";
            }
            button.LongDescription = string.Format("Верстия: {0}\nСборка: {1}-{2}", ModuleData.Version, ModuleData.Build, ModuleData.Date);
            button.ItemText = text;
            button.LargeImage = new BitmapImage(new Uri(imageSource.Value));
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=672"));
        }
    }
}
