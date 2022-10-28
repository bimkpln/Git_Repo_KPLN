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
using System.IO;

namespace KPLN_ModelChecker_User
{
    public class Module : IExternalModule
    {
        private string _mainContextualHelp = "http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=937";
        private readonly string _AssemblyPath = Assembly.GetExecutingAssembly().Location;

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
            string assembly = _AssemblyPath.Split(new string[] { "\\" }, StringSplitOptions.None).Last().Split('.').First();
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Контроль качества");
            PulldownButtonData pullDownData = new PulldownButtonData("Проверить", "Проверить");
            pullDownData.ToolTip = "Набор плагинов, для ручной проверки моделей на ошибки";
            pullDownData.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, _mainContextualHelp));

            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            BtnImagine(pullDown, "checker_push.png");
            
            AddPushButtonData(
                "CheckLevels", 
                "Проверка\nуровней", 
                "Проверить все элементы в проекте на правильность расположения относительно связанного уровня.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckLevelOfInstances).FullName,
                pullDown,
                "checker_levels.png",
                "https://clck.ru/32GG7V"
                );
            AddPushButtonData(
                "CheckMirrored", 
                "Проверка\nзеркальных", 
                "Проверка проекта на наличие зеркальных элементов (<Окна>, <Двери>).",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckMirroredInstances).FullName,
                pullDown,
                "checker_mirrored.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckCoordinates", 
                "Проверка\nсвязей", 
                "Проверка подгруженных rvt-связей:" +
                "\n1. Корректность настройки общей площадки Revit;" +
                "\n2. Корректность заданного рабочего набора;" +
                "\n3. Прикрепление экземпляра связи.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandLinks).FullName,
                pullDown,
                "checker_locations.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckLevelMonitored", 
                "Мониторинг\nуровней", "Проверка элементов на наличие настроенного мониторинга, а также на наличие прикрепления.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckLevels).FullName,
                pullDown,
                "checker_levels_monitor.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckGridMonitored",
                "Мониторинг\nосей",
                "Проверка элементов на наличие настроенного мониторинга, а также на наличие прикрепления.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckGrids).FullName,
                pullDown,
                "checker_grids_monitor.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckNames", 
                "Проверка\nимен", 
                "Проверка семейств и типоразмеров на наличие дубликатов имен.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckNames).FullName,
                pullDown,
                "family_name.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckWorksets", 
                "Проверка\nрабочих наборов", 
                "Проверка элементов на корректность рабочих наборов.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckElementWorksets).FullName,
                pullDown,
                "checker_worksets.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckDimensions",
                "Проверка размеров",
                "Анализирует все размеры, на предмет:" +
                    "\n1. Замены значения;" +
                    "\n2. Округления значений размеров с нарушением требований пункта 5.1 ВЕР.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckDimensions).FullName,
                pullDown,
                "dimensions.png",
                _mainContextualHelp
                );
            AddPushButtonData(
                "CheckAnnotations", 
                "Проверка листов на аннотации", 
                "Анализирует все элементы на листах и ищет аннотации следующих типов:" +
                    "\n1. Линии детализации;" + 
                    "\n2. Элементы узлов;" + 
                    "\n3. Текст;" +
                    "\n4. Типовые аннотации;" +
                    "\n5. Изображения.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckListAnnotations).FullName,
                pullDown,
                "surch_list_annotation.png",
                _mainContextualHelp
                );
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

        /// <summary>
        /// Метод для добавления отдельной в панель
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="panel">Панель, в которую добавляем кнопку</param>
        /// <param name="imageName">Имя иконки</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonData(string name, string text, string description, string longDescription, string className, PulldownButton pullDown, string imageName, string anchorlHelp)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = pullDown.AddPushButton(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, anchorlHelp));
            BtnImagine(button, imageName);
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="button">Кнопка, куда нужно добавить иконку</param>
        /// <param name="imageName">Имя иконки с раширением</param>
        private void BtnImagine(RibbonButton button, string imageName)
        {
            string imageFullPath = Path.Combine(new FileInfo(_AssemblyPath).DirectoryName, @"Source\", imageName);
            button.LargeImage = new BitmapImage(new Uri(imageFullPath));
        }
    }
}
