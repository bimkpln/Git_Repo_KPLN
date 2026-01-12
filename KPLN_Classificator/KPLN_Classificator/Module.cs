using Autodesk.Revit.UI;
using KPLN_Classificator.ExecutableCommand;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using static KPLN_Classificator.ApplicationConfig;

namespace KPLN_Classificator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Module()
        {
            //Конфигурирование приложения для работы в среде KPLN_Loader
            CurrentOutput = new KplnOutput();
            CurrentCmdEnv = new KplnCommandEnvironment();
        }

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            try { application.CreateRibbonTab(tabName); } catch { }

            string panelName = "Параметры";
            RibbonPanel panel = null;
            List<RibbonPanel> tryPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName).ToList();
            if (tryPanels.Count == 0)
                panel = application.CreateRibbonPanel(tabName, panelName);
            else
                panel = tryPanels.First();

            AddPushButtonDataInPanel(
                "ClassificatorCompleteCommand",
                "Заполнить\nпараметры",
                "Параметризация элементов согласно заданным правилам.",
                "Возможности:\n" +
                    "1. Задание правил для параметризации элементов;\n" +
                    "2. Маппинг параметров (передача значений между параметрами элемента);\n" +
                    "3. Сохранение конфигурационного файла с возможностью повторного использования.\n" +
                    $"\nДата сборки: {Date}\nНомер сборки: {Version}\nИмя модуля: {ModuleName}",
                typeof(CommandOpenClassificatorForm).FullName,
                panel,
                "classificator",
                "http://moodle.stinproject.local"
            );


            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для добавления отдельной кнопки в панель
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="panel">Панель, в которую добавляем кнопку</param>
        /// <param name="imageName">Имя иконки, как ресурса</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            button.AvailabilityClassName = typeof(Availability.StaticAvailable).FullName;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }
    }
}
