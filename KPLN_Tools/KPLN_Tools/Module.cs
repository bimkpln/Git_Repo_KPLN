using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using static KPLN_Tools.ModuleData;
using static KPLN_Loader.Output.Output;
using KPLN_Tools.Tools;

namespace KPLN_Tools
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
                MainWindowHandle = Tools.WindowHandleSearch.MainWindowHandle.Handle;
            }
            catch (Exception e)
            {
                PrintError(e);
            }
#endif
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Инструменты");
            PulldownButtonData pullDownData = new PulldownButtonData("Инструменты", "Инструменты");
            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            pullDown.LargeImage = new BitmapImage(new Uri(new Source.Source(Common.Collections.Icon.toolBox).Value));
            string assembly = Assembly.GetExecutingAssembly().Location.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Split('.').First();
            AddPushButtonData("Перенумеровать", "Перенумеровать листы", "Перенумеровать листы по заданной функции", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandRenamer"), pullDown, new Source.Source(Common.Collections.Icon.renamerFunc), "http://moodle.stinproject.local");
            AddPushButtonData("Нумерация", "Нумерация", "Нумерация позици в спецификации на +1 от начального значения", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandAutonumber"), pullDown, new Source.Source(Common.Collections.Icon.autonumber), "http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=687");
            
            return Result.Succeeded;
        }
        /// <summary>
        /// Стандартный метод который копируется из проекта в проект для быстрого созданий кнопки
        /// </summary>
        /// <param name="name"></param>
        /// <param name="text"></param>
        /// <param name="description"></param>
        /// <param name="className"></param>
        /// <param name="pullDown"></param>
        /// <param name="imageSource"></param>
        private void AddPushButtonData(string name, string text, string description, string className, PulldownButton pullDown, Source.Source imageSource, string manualPage)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = pullDown.AddPushButton(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = string.Format("Версия: {0}\nСборка: {1}-{2}", ModuleData.Version, ModuleData.Build, ModuleData.Date);
            button.ItemText = text;
            //Ссылка на web-страницу по клавише F1
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, manualPage));
            button.LargeImage = new BitmapImage(new Uri(imageSource.Value));
        }
    }
}
