using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ParamSetter.Tools;
using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using static KPLN_Loader.Output.Output;
using static KPLN_ParamSetter.ModuleData;

namespace KPLN_ParamSetter
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
            catch (Exception e)
            {
                PrintError(e);
            }
#endif
            string assembly = Assembly.GetExecutingAssembly().Location.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Split('.').First();
            string ribbonName = "Параметры";
            RibbonPanel panel = application.CreateRibbonPanel(tabName, ribbonName);
            AddPushButtonData("Открыть окно переноса параметров", "Перенести\nпараметры", "Копирование значений из параметра в параметр", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandOpenSetManager"), panel, new Source.Source(Common.Collections.Icon.ParamSetter));
            return Result.Succeeded;
        }
        private void AddPushButtonData(string name, string text, string description, string className, RibbonPanel panel, Source.Source imageSource)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = string.Format("Верстия: {0}\nСборка: {1}-{2}", ModuleData.Version, ModuleData.Build, ModuleData.Date);
            button.ItemText = text;
            button.LargeImage = new BitmapImage(new Uri(imageSource.Value));
        }
    }
}
