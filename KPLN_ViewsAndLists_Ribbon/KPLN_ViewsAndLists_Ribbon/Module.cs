using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ViewsAndLists_Ribbon.ExternalCommands.Lists;
using KPLN_ViewsAndLists_Ribbon.ExternalCommands.Views;
using System.Reflection;

namespace KPLN_ViewsAndLists_Ribbon
{
    public class Module : IExternalModule
    {
        public readonly static string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);


            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Виды и листы");

            // Добавляю выпадающие списки pullDown для видов
            PulldownButtonData pullDownData_Views = new PulldownButtonData("Views", "Виды")
            {
                ToolTip = "Пакетная работа с видами"
            };
            PulldownButton pullDown_Views = panel.AddItem(pullDownData_Views) as PulldownButton;
            pullDown_Views.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainViews", 32);
            pullDown_Views.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainViews", 32);
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((pullDown_Views, "mainViews", _assemblyName));
#endif


            // Добавляю выпадающие списки pullDown для листов
            PulldownButtonData pullDownData_Lists = new PulldownButtonData("Lists", "Листы")
            {
                ToolTip = "Пакетная работа с листами"
            };
            PulldownButton pullDown_Lists = panel.AddItem(pullDownData_Lists) as PulldownButton;
            pullDown_Lists.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainLists", 32);
            pullDown_Lists.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainLists", 32);
#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((pullDown_Lists, "mainLists", _assemblyName));
#endif

            #region Добавляю в выпадающий список элементы для видов
            AddPushButtonDataInPullDown(
                ExtCmdCutCopy.PluginName,
                ExtCmdCutCopy.PluginName,
                "Копирует подрезку активного плана на выбранные из списка планы этажей/потолков/несущих конструкций.",
                string.Format(
                    "Запусти на плане-доноре, и выбери из списка другие планы, куда нужно скопировать подрезку\n" +
                    "Помимо копирования границ подрезки, плагин копирует и настройки видимости подрезки." +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmdCutCopy).FullName,
                pullDown_Views,
                "CutCopy",
                "http://moodle/mod/book/view.php?id=502&chapterid=1295"
            );

            AddPushButtonDataInPullDown(
                "ViewTemplateCopy",
                "Пакетное копирование\nшаблонов вида",
                "Копировать шаблоны вида из другого проекта",
                string.Format(
                    "Пакетное копирование шаблонов видов из одного проекта в один или несколько других файлов без необходимости их открытия." +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmdViewTemplateCopy).FullName,
                pullDown_Views,
                "ViewTemplateCopy",
                "http://moodle/mod/book/view.php?id=502&chapterid=1339"
            );

            AddPushButtonDataInPullDown(
                "BatchCreate",
                "Создать\nфильтры",
                "Создать фильтры",
                string.Format(
                    "Создает набор фильтров графики по критериям из CSV-файла. Примеры файлов в папке с программой.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Views.ExtCmdCreate).FullName,
                pullDown_Views,
                "CommandCreate",
                "http://moodle/mod/book/view.php?id=502&chapterid=670"
            );

            AddPushButtonDataInPullDown(
                "BatchDelete",
                "Удалить\nфильтры",
                "Удалить фильтры",
                string.Format(
                    "Пакетное удаления фильтров графики.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Views.ExtCmdBatchDelete).FullName,
                pullDown_Views,
                "CommandBatchDelete",
                "http://moodle/mod/book/view.php?id=502&chapterid=670"
            );

            AddPushButtonDataInPullDown(
                "ViewColoring",
                "Раскрасить",
                "Колоризация\nвида",
                string.Format(
                    "Создает фильтры и раскрашивает элементы разными цветами в зависимости от значения параметра.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Views.ExtCmdViewColoring).FullName,
                pullDown_Views,
                "CommandViewColoring",
                "http://moodle/mod/book/view.php?id=502&chapterid=671l"
            );


            AddPushButtonDataInPullDown(
                "WallHatch",
                "Штриховки\nстен",
                "Штриховка по высоте стен",
                string.Format(
                    "Вычисляет отметки верха и низа стен;\n" +
                    "Создает набор фильтров и выделяет стены разными штриховками;\n" +
                    "Записывает условное обозначение, соответствующее штриховке.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Views.ExtCmdWallHatch).FullName,
                pullDown_Views,
                "CommandWallHatch",
                "http://bim-starter.com/plugins/wallhatch/"
            );
            #endregion

            #region Добавляю в выпадающий список элементы для листов
            AddPushButtonDataInPullDown(
                ExtCmdListRenumber.PluginName,
                ExtCmdListRenumber.PluginName,
                "Перенумеровать листы",
                string.Format(
                    "Изменяет нумерацию по заданной функции;\n" +
                    "Заполняет выбранный символ Юникода.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmdListRenumber).FullName,
                pullDown_Lists,
                "CommandListRename",
                "http://moodle/mod/book/view.php?id=502&chapterid=911"
            );

            AddPushButtonDataInPullDown(
                "CommandListTBlockParamCopier",
                "Перенос\nпараметров",
                "Перенос параметров листа в параметры основной надписи",
                string.Format(
                    "Переносит значения листа в экземпляр основной надписи.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Lists.ExtCmdListTBlockParamCopier).FullName,
                pullDown_Lists,
                "CommandListTBlockParamCopier",
                "http://moodle/mod/book/view.php?id=502&chapterid=911"
            );

            AddPushButtonDataInPullDown(
                "CommandListRevisionClouds",
                "Изменения и Пометочные облака",
                "Плагин автоматически выполняет на всех листах следующее:\n" +
                "1. Заполняет ячейку \"Кол.уч\" в штампе по количеству пометочных облаков\n" +
                "2. Перечисляет изменения на листе в параметре \"Ш.ПримечаниеЛиста\" в формате: \"Изм. 1(Зам.), Изм. 2(Нов.)...\"— для отображения в спецификации\n" +
                "\"О_Ведомость рабочих чертежей основного комплекта...\"",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Lists.ExtCmdListRevisionClouds).FullName,
                pullDown_Lists,
                "CommandListRevisionClouds",
                "http://moodle/mod/book/view.php?id=502&chapterid=1330"
            );
            #endregion

            return Result.Succeeded;
        }

        public Result Close() => Result.Succeeded;

        /// <summary>
        /// Метод для добавления кнопки в выпадающий список
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="description">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="pullDown">Выпадающий список, в который добавляем кнопку</param>
        /// <param name="imageName">Имя иконки</param>
        /// <param name="anchorlHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPullDown(string name, string text, string description, string longDescription, string className, PulldownButton pullDown, string imageName, string anchorlHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = pullDown.AddPushButton(data) as PushButton;
            button.ToolTip = description;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, anchorlHelp));
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }
    }
}
