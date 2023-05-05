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
using System.Windows.Media;

namespace KPLN_ModelChecker_User
{
    public class Module : IExternalModule
    {
        private string _mainContextualHelp = "http://moodle/mod/book/view.php?id=502&chapterid=937";
        private int _userDepartment = KPLN_Loader.Preferences.User.Department.Id;
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
            pullDownData.Image = PngImageSource("KPLN_ModelChecker_User.Source.checker_push.png");
            pullDownData.LargeImage = PngImageSource("KPLN_ModelChecker_User.Source.checker_push.png");
            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            
            AddPushButtonData(
                "CheckLevels", 
                "Проверка\nуровней", 
                "Проверить все элементы в проекте на правильность расположения относительно связанного уровня.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckLevelOfInstances).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_levels.png",
                _mainContextualHelp,
                _userDepartment != 3
                );
            AddPushButtonData(
                "CheckMirrored", 
                "Проверка\nзеркальных", 
                "Проверка проекта на наличие зеркальных элементов (<Окна>, <Двери>).",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckMirroredInstances).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_mirrored.png",
                _mainContextualHelp,
                _userDepartment != 2
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
                "KPLN_ModelChecker_User.Source.checker_locations.png",
                _mainContextualHelp,
                true
                );
            AddPushButtonData(
                "CheckLevelMonitored", 
                "Мониторинг\nуровней", "Проверка элементов на наличие настроенного мониторинга, а также на наличие прикрепления.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckLevels).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_levels_monitor.png",
                _mainContextualHelp,
                true
                );
            AddPushButtonData(
                "CheckGridMonitored",
                "Мониторинг\nосей",
                "Проверка элементов на наличие настроенного мониторинга, а также на наличие прикрепления.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckGrids).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_grids_monitor.png",
                _mainContextualHelp,
                true
                );
            AddPushButtonData(
                "CheckNames", 
                "Проверка\nсемейств",
                "Проверка семейств на:" +
                    "\n1. Импорт семейств из разрешенных источников (диск Х);" +
                    "\n2. На наличие дубликатов имен (проверяются и типоразмеры).",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckFamilies).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.family_name.png",
                _mainContextualHelp,
                true
                );
            AddPushButtonData(
                "CheckWorksets", 
                "Проверка\nрабочих наборов", 
                "Проверка элементов на корректность рабочих наборов.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckElementWorksets).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_worksets.png",
                _mainContextualHelp,
                true
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
                "KPLN_ModelChecker_User.Source.dimensions.png",
                _mainContextualHelp,
                true
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
                "KPLN_ModelChecker_User.Source.surch_list_annotation.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=991#:~:text=%D0%92%D0%AB%D0%9F%D0%90%D0%94%D0%90%D0%AE%D0%A9%D0%98%D0%99%20%D0%A1%D0%9F%D0%98%D0%A1%D0%9E%D0%9A%20%22%D0%9F%D0%A0%D0%9E%D0%92%D0%95%D0%A0%D0%98%D0%A2%D0%AC%22-,%D0%9F%D0%A0%D0%9E%D0%92%D0%95%D0%A0%D0%9A%D0%98%20%D0%9B%D0%98%D0%A1%D0%A2%D0%9E%D0%92%20%D0%9D%D0%90%20%D0%90%D0%9D%D0%9D%D0%9E%D0%A2%D0%90%D0%A6%D0%98%D0%98,-%D0%A2%D0%B5%D0%B3%D0%B8%3A",
                true
                );
            AddPushButtonData(
                "CheckFlatsArea",
                "Проверка площадей квартир",
                "Сравнить фактические значения площадей (по квартирографии) со значениями, зафиксированными на стадии П (после выхода из экспертизы):" +
                    "\n1. Находит разницу имен и номеров помещений;" +
                    "\n2. Находит разницу в суммарной площади (физической) квартиры, если она превышает 1 м²;" +
                    "\n3. Находит разницу в площади помещения вне квартиры, если она превышает 1 м²;" +
                    "\n4. Находит разницу в значениях параметров площадей в марках и фактической, если она превышает 0,1 м²;" +
                    "\n5. Находит разницу зафиксированной площади квартиры.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(ExternalCommands.CommandCheckFlatsArea).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_flatsArea.png",
                _mainContextualHelp,
                _userDepartment != 2 && _userDepartment != 3
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
        private void AddPushButtonData(string name, string text, string description, string longDescription, string className, PulldownButton pullDown, string imageName, string anchorlHelp, bool isVisible)
        {
            PushButtonData data = new PushButtonData(name, text, Assembly.GetExecutingAssembly().Location, className);
            PushButton button = pullDown.AddPushButton(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, anchorlHelp));
            button.Image = PngImageSource(imageName);
            button.LargeImage = PngImageSource(imageName);
            button.Visible = isVisible;
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="button">Кнопка, куда нужно добавить иконку</param>
        /// <param name="imageName">Имя иконки с раширением</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
