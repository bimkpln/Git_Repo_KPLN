using Autodesk.Revit.UI;
using KPLN_Classificator.ExecutableCommand;
using KPLN_Library_Forms.UI.HtmlWindow;
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


            //Ищу  выпадающий список
            string pullDownName = "Параметры";
            IList<RibbonItem> tryParamsPullDownBtns = panel.GetItems();
            if (!(tryParamsPullDownBtns.FirstOrDefault(ri => ri.Name.Equals(pullDownName)) is PulldownButton paramsPullDownBtn))
            {
                HtmlOutput.Print(
                    "Отправь разработчику - ошибка инициализации плагина KPLN_Classificator. " +
                        "Нет выпадающего списка для кнопки. Нарушен порядок плагинов в БД",
                    MessageType.Error);

                return Result.Cancelled;
            }


            AddPushButtonDataInPullDown(
                "ClassificatorCompleteCommand",
                "Заполнить\nпараметры",
                "Параметризация элементов согласно заданным правилам.",
                "Возможности:\n" +
                    "1. Задание правил для параметризации элементов;\n" +
                    "2. Маппинг параметров (передача значений между параметрами элемента);\n" +
                    "3. Сохранение конфигурационного файла с возможностью повторного использования.\n" +
                    $"\nДата сборки: {Date}\nНомер сборки: {Version}\nИмя модуля: {ModuleName}",
                typeof(CommandOpenClassificatorForm).FullName,
                paramsPullDownBtn,
                "classificator",
                "http://moodle.stinproject.local"
            );


            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для создания PushButtonData будущей кнопки
        /// </summary>
        private void AddPushButtonDataInPullDown(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            PulldownButton pullDownButton,
            string imageName,
            string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = pullDownButton.AddPushButton(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16);
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }
    }
}
