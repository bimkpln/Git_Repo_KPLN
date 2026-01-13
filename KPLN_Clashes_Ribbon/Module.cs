using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KPLN_Clashes_Ribbon
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            //Ищу или создаю панель
            string panelName = "Менеджеры";
            RibbonPanel panel = null;
            IEnumerable<RibbonPanel> tryMPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryMPanels.Any())
                panel = tryMPanels.FirstOrDefault();
            else
                panel = application.CreateRibbonPanel(tabName, panelName);

            //Добавляю кнопку в панель (тут приведен пример поиска панели, вместо этого - панель можно создать)
            AddPushButtonDataInPanel(
                "NWC проверки",
                "Менеджер\nпроверок",
                "Утилита для работы с отчетами Navisworks.",
                string.Format(
                    "Осуществляет поиск коллизий по координатам из Navisworks.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(Commands.CommandShowManager).FullName,
                panel,
                "icon",
                "http://moodle/mod/book/view.php?id=502&chapterid=672",
                true
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
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string contextualHelp, bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

            if (avclass) button.AvailabilityClassName = typeof(Availability.StaticAvailable).FullName;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }
    }
}
