using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Library_SQLiteWorker.FactoryParts;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_Lib.Common;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExternalCommands;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ModelChecker_User
{
    public class Module : IExternalModule
    {
        private readonly string _mainContextualHelp = "http://moodle/mod/book/view.php?id=502&chapterid=991";
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Module()
        {
            UserDbService userDbService = (UserDbService)new CreatorUserDbService().CreateService();
            CurrentDbUser = userDbService.GetCurrentDBUser();
        }

        /// <summary>
        /// Текущий пользователь
        /// </summary>
        public static DBUser CurrentDbUser { get; private set; }

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            #region Инициализация элементов нужно для плагина проверки факта запуска
            // Инициирую статические поля проверок
            CommandCheckDimensions commandCheckDimensions = new CommandCheckDimensions(new ExtensibleStorageEntity(
                CommandCheckDimensions.PluginName,
                "KPLN_CheckDimensions",
                new Guid("f2e615e0-a15b-43df-a199-a88d18a2f568"),
                new Guid("f2e615e0-a15b-43df-a199-a88d18a2f569")));
            CheckWorksets checkWorksets = new CheckWorksets();
            CommandCheckFamilies commandCheckFamilies = new CommandCheckFamilies(new ExtensibleStorageEntity(
                CommandCheckFamilies.PluginName,
                "KPLN_CommandCheckFamilies",
                new Guid("168c83b9-1d62-4d3f-9bbb-fd1c1e9a0807"),
                new Guid("168c83b9-1d62-4d3f-9bbb-fd1c1e9a0808")));
            CheckMainLines checkMainLines = new CheckMainLines();
            CommandCheckFlatsArea commandCheckFlatsArea = new CommandCheckFlatsArea(new ExtensibleStorageEntity(
                CommandCheckFlatsArea.PluginName,
                "KPLN_CheckFlatsArea",
                new Guid("720080C5-DA99-40D7-9445-E53F288AA150"),
                new Guid("720080C5-DA99-40D7-9445-E53F288AA151"),
                new Guid("720080C5-DA99-40D7-9445-E53F288AA155")));
            CommandCheckHoles commandCheckHoles = new CommandCheckHoles(new ExtensibleStorageEntity(
                CommandCheckHoles.PluginName,
                "KPLN_CheckHoles",
                new Guid("820080C5-DA99-40D7-9445-E53F288AA160"),
                new Guid("820080C5-DA99-40D7-9445-E53F288AA161")));
            CommandCheckLevelOfInstances сommandCheckLevelOfInstances = new CommandCheckLevelOfInstances(new ExtensibleStorageEntity(
                CommandCheckLevelOfInstances.PluginName,
                "KPLN_CheckLevelOfInstances",
                new Guid("bb59ea6c-9208-4fae-b609-3d73dc3abf52"),
                new Guid("bb59ea6c-9208-4fae-b609-3d73dc3abf53")));
            CheckLinks checkLinks = new CheckLinks();
            CommandCheckListAnnotations commandCheckListAnnotations = new CommandCheckListAnnotations(new ExtensibleStorageEntity(
                CommandCheckListAnnotations.PluginName,
                "KPLN_CheckAnnotation",
                new Guid("caf1c9b7-14cc-4ba1-8336-aa4b357d2898")));
            CommandCheckMEPHeight commandCheckMEPHeight = new CommandCheckMEPHeight(new ExtensibleStorageEntity(
                CommandCheckMEPHeight.PluginName,
                "KPLN_CheckMEPHeight",
                new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b68"),
                new Guid("1c2d57de-4b61-4d2b-a81b-070d5aa76b69")));
            CommandCheckMirroredInstances commandCheckMirroredInstances = new CommandCheckMirroredInstances(new ExtensibleStorageEntity(
                CommandCheckMirroredInstances.PluginName,
                "KPLN_CheckMirroredInstances",
                new Guid("33b660af-95b8-4d7c-ac42-c9425320557b"),
                new Guid("33b660af-95b8-4d7c-ac42-c9425320557c")));

            // Запись в массив для передачи ExtensibleStorageEntity в CommandCheckLaunchDate (после инициализации статических полей)
            ExtensibleStorageEntity[] extensibleStorageEntities = new ExtensibleStorageEntity[]
            {
                // Проверки из этой сборки
                CommandCheckDimensions.ESEntity,
                checkWorksets.ESEntity,
                CommandCheckFamilies.ESEntity,
                checkMainLines.ESEntity,
                CommandCheckFlatsArea.ESEntity,
                CommandCheckHoles.ESEntity,
                CommandCheckLevelOfInstances.ESEntity,
                checkLinks.ESEntity,
                CommandCheckMEPHeight.ESEntity,
                CommandCheckMirroredInstances.ESEntity,
                // Сторонние плагины (добавлять из исходников)
                new ExtensibleStorageEntity("АР_П: Фиксация площадей", "KPLN_ARArea", new Guid("720080C5-DA99-40D7-9445-E53F288AA155")),
                new ExtensibleStorageEntity("ОВ: Толщина воздуховодов", "KPLN_DuctSize", new Guid("753380C4-DF00-40F8-9745-D53F328AC139")),
                new ExtensibleStorageEntity("ОВВК: Спецификации", "KPLN_IOSQuant", new Guid("720080C5-DA99-40D7-9445-E53F288AA140")),
                new ExtensibleStorageEntity("ОВ: Клапаны ОЗК", "KPLN_OZKAccessory", new Guid("85c46d4e-fdc6-424e-909a-27af56597328")),
                new ExtensibleStorageEntity("ИОС: Имя системы", "KPLN_SystemType", new Guid("be15305c-5249-4581-a4ca-01784efd8415")),
            };
            CommandCheckLaunchDate commandCheckLaunchDate = new CommandCheckLaunchDate(extensibleStorageEntities);
            #endregion

            //Добавляю кнопку в панель
            string currentPanelName = "Контроль качества";
            RibbonPanel currentPanel = application.GetRibbonPanels(tabName).Where(i => i.Name == currentPanelName).ToList().FirstOrDefault() ?? application.CreateRibbonPanel(tabName, "Контроль качества");

            PulldownButtonData pullDownData = new PulldownButtonData("Проверить", "Проверить")
            {
                ToolTip = "Набор плагинов, для ручной проверки моделей на ошибки"
            };
            pullDownData.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, _mainContextualHelp));
            pullDownData.Image = PngImageSource("KPLN_ModelChecker_User.Source.checker_push.png");
            pullDownData.LargeImage = PngImageSource("KPLN_ModelChecker_User.Source.checker_push.png");
            PulldownButton pullDown = currentPanel.AddItem(pullDownData) as PulldownButton;

            AddPushButtonData(
                "CheckLaunchDate",
                CommandCheckLaunchDate.PluginName,
                "Проверить факт и дату запуска плагинов.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckLaunchDate).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.launchDate.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckCoordinates",
                checkLinks.PluginName,
                "Проверка подгруженных rvt-связей:" +
                "\n1. Корректность настройки общей площадки Revit;" +
                "\n2. Корректность заданного рабочего набора;" +
                "\n3. Прикрепление экземпляра связи.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckLinks).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_locations.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckMainLines",
                checkMainLines.PluginName,
                "Анализирует оси и уровни на:" +
                    "\n1. Наличие и корректность мониторинга;" +
                    "\n2. Наличие прикрепления;" +
                    "\n3. Точность расстояния параллельных ПРЯМЫХ осей.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckMainLines).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_mainLines.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckNames",
                CommandCheckFamilies.PluginName,
                "Проверка семейств на:" +
                    "\n1. Импорт семейств из разрешенных источников (диск Х);" +
                    "\n2. Наличие дубликатов имен (проверяются и типоразмеры);" +
                    "\n3. Соответсвие имени типоразмеров системных семейств требованиям ВЕР.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckFamilies).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.family_name.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckWorksets",
                checkWorksets.PluginName,
                "Проверка элементов на корректность следующих рабочих наборов:"+
                    "\n1. РН для связей;" +
                    "\n2. РН для осей и уровней;" +
                    "\n3. РН для скопированных и замониторенных элементов из других моделей.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckWorksets).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_worksets.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckDimensions",
                CommandCheckDimensions.PluginName,
                "Анализирует все размеры, на предмет:" +
                    "\n1. Замены значения;" +
                    "\n2. Округления значений размеров с нарушением требований пункта 5.1 ВЕР.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckDimensions).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.dimensions.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckAnnotations",
                CommandCheckListAnnotations.PluginName,
                "Анализирует все элементы на листах и ищет аннотации следующих типов:" +
                    "\n1. Линии детализации;" +
                    "\n2. Элементы узлов;" +
                    "\n3. Текст;" +
                    "\n4. Типовые аннотации;" +
                    "\n5. Изображения.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckListAnnotations).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.surch_list_annotation.png",
                _mainContextualHelp,
                true
                );

            AddPushButtonData(
                "CheckMonolith",
                CommandCheckMonolith.PluginName,
                "Плагин предназначен для автоматического сравнения  внешних контуров перекрытий ипилонов из основной модели и связанных моделей на основе геометрических координат и объёма бетона.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckMonolith).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_monolith.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckLevels",
                CommandCheckLevelOfInstances.PluginName,
                "Проверить все элементы в проекте на правильность расположения относительно связанного уровня.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckLevelOfInstances).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_levels.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 2 || CurrentDbUser.SubDepartmentId == 3 || CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckMirrored",
                CommandCheckMirroredInstances.PluginName,
                "Проверка проекта на наличие зеркальных элементов (<Окна>, <Двери>).",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckMirroredInstances).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_mirrored.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 2 || CurrentDbUser.SubDepartmentId == 4 || CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckWetZones",
                CommandCheckWetZones.PluginName,
                "Плагин, который позволяет находить в проектах несоответствия СП 54.13330.2022, п 7.20, 7.21.\n",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckWetZones).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_wet_zones.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 2 || CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckHoles",
                CommandCheckHoles.PluginName,
                "Плагин выполняет следующие функции:\n" +
                        "1. Проверяет отверстия, в которых нет лючков на наличие в нем элементов ИОС;\n" +
                        "2. Проверяет отверстия, в которых нет лючков на заполненность элементами ИОС.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckHoles).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checkHoles.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 2 || CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckFlatsArea",
                CommandCheckFlatsArea.PluginName,
                "Стадия РД - сравнить фактические значения площадей (по квартирографии) со значениями, зафиксированными на стадии П (после выхода из экспертизы):" +
                    "\n1. Находит разницу имен и номеров помещений;" +
                    "\n2. Находит разницу в суммарной площади (физической) квартиры, если она превышает 1 м²;" +
                    "\n3. Находит разницу в площади помещения вне квартиры, если она превышает 1 м²;" +
                    "\n4. Находит разницу в значениях параметров площадей в марках и фактической, если она превышает 0,1 м²;" +
                    "\n5. Находит разницу зафиксированной площади квартиры.",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckFlatsArea).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_flatsArea.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 2 || CurrentDbUser.SubDepartmentId == 8
                );

            AddPushButtonData(
                "CheckMEPHeight",
                CommandCheckMEPHeight.PluginName,
                "Найти элементы, которые расположены в границах помещений на высоте, меньше 2.2 м:",
                $"\nДата сборки: {ModuleData.Date}\nНомер сборки: {ModuleData.Version}\nИмя модуля: {ModuleData.ModuleName}",
                typeof(CommandCheckMEPHeight).FullName,
                pullDown,
                "KPLN_ModelChecker_User.Source.checker_mepHeigtheight.png",
                _mainContextualHelp,
                CurrentDbUser.SubDepartmentId == 8
                );

            return Result.Succeeded;
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
        private void AddPushButtonData(string name, string text, string description, string longDescription, string className, PulldownButton pullDown, string imageName, string anchorlHelp, bool isVisible)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
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
        /// <param name="embeddedPathname">Имя иконки с раширением</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
