using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using static KPLN_ModelChecker_User.ModuleData;
using static KPLN_Loader.Output.Output;
using System.Windows.Interop;
using System.Windows;
using KPLN_ModelChecker_User.Tools;

namespace KPLN_ModelChecker_User
{
    public class Module : IExternalModule
    {
        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
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
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Контроль качества");
            PulldownButtonData pullDownData = new PulldownButtonData("Проверки", "Проверки");
            pullDownData.ToolTip = "Набор плагинов, для ручной проверки моделей на ошибки";

            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            pullDown.LargeImage = new BitmapImage(new Uri(new Source.Source(Common.Collections.Icon.Errors).Value));
            
            AddPushButtonData("Проверить уровни элементов", "Проверка\nуровней", "Проверить все элементы в проекте на правильность расположения относительно связанного уровня.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckLevelOfInstances"), pullDown, new Source.Source(Common.Collections.Icon.CheckLevels));
            AddPushButtonData("Найти зеркальные элементы", "Проверка\nзеркальных", "Проверка проекта на наличие зеркальных элементов (<Окна>, <Двери>).", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckMirroredInstances"), pullDown, new Source.Source(Common.Collections.Icon.CheckMirrored));
            AddPushButtonData("Проверить площадки", "Проверка\nсвязей", "Проверка подгруженных rvt-связей на правильность настройки общей площадки Revit и выбранного рабочего набора.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckPosition"), pullDown, new Source.Source(Common.Collections.Icon.CheckLocations));
            AddPushButtonData("Проверить мониторинг уровней", "Проверка\nуровней", "Проверка элементов на наличие настроенного мониторинга.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckMonitoringLevels"), pullDown, new Source.Source(Common.Collections.Icon.CheckMonitorLevels));
            AddPushButtonData("Проверить мониторинг осей", "Проверка\nосей", "Проверка элементов на наличие настроенного мониторинга.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckMonitoringGrids"), pullDown, new Source.Source(Common.Collections.Icon.CheckMonitorGrids));
            AddPushButtonData("Проверить имена семейств", "Проверка\nимен", "Проверка семейств и типоразмеров на наличие дубликатов имен.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckNames"), pullDown, new Source.Source(Common.Collections.Icon.FamilyName));
            AddPushButtonData("Проверить принадлежность рабочим наборам", "Проверка\nрабочих наборов", "Проверка элементов на корректность рабочих наборов.", string.Format("{0}.{1}", assembly, "ExternalCommands.CommandCheckElementWorksets"), pullDown, new Source.Source(Common.Collections.Icon.Worksets));
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            return Result.Succeeded;
        }
        private void OnIdling(object sender, IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIControlledApplication controlledApplication = sender as UIControlledApplication;
            while (CommandQueue.Count != 0)
            {
                using (Transaction t = new Transaction(uiapp.ActiveUIDocument.Document, ModuleName))
                {
                    t.Start();
                    try
                    {
                        Result result = CommandQueue.Dequeue().Execute(uiapp);
                        if (result != Result.Succeeded)
                        {
                            t.RollBack();
                            break;
                        }
                        else
                        {
                            t.Commit();
                        }
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                        t.RollBack();
                        break;
                    }
                }
            }
        }
        private void AddPushButtonData(string name, string text, string description, string className, PulldownButton pullDown, Source.Source imageSource)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = pullDown.AddPushButton(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = string.Format("Верстия: {0}\nСборка: {1}-{2}", ModuleData.Version, ModuleData.Build, ModuleData.Date);
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, ModuleData.ManualPage));
            button.LargeImage = new BitmapImage(new Uri(imageSource.Value));
        }
    }
}
