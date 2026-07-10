using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Parameters_Ribbon.Common;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KPLN_Parameters_Ribbon
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            //Ищу или создаю панель
            string panelName = "Параметры";
            RibbonPanel panel = null;
            IEnumerable<RibbonPanel> tryMPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryMPanels.Any())
                panel = tryMPanels.FirstOrDefault();
            else
                panel = application.CreateRibbonPanel(tabName, panelName);


            //Добавляю выпадающий список pullDown
            PulldownButtonData uparamsPullDownBtnData = new PulldownButtonData("Параметры", "Параметры")
            {
                ToolTip = "Коллекция плагинов по работе с параметрами в проекте",
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "mainParams", 32),
            };
            PulldownButton paramsPullDownBtn = panel.AddItem(uparamsPullDownBtnData) as PulldownButton;
            
            SetRIShowText(paramsPullDownBtn, false);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((paramsPullDownBtn, "mainParams", _assemblyName));
#endif


            # region Добавляю кнопки в панель
            AddPushButtonDataInPullDown(
                "Копирование параметров проекта",
                "Параметры\nпроекта",
                "Производит копирование параметров проекта из файла, содержащего сведения о проекте.\n" +
                    "Для копирования параметров необходимо открыть исходный файл или подгрузить его как связь.",
                string.Format(
                    "Алгоритм запуска:\n" +
                        "1. Выделяем стартовую ячейку спецификации;\n" +
                        "2. Вводим данные, которые указаны в окне.\n\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandCopyProjectParams).FullName,
                paramsPullDownBtn,
                "copyProjectParams",
                "http://moodle/mod/book/view.php?id=502&chapterid=992#:~:text=%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%D0%90%22-,%D0%9F%D0%A3%D0%A2%D0%AC,-%D0%9F%D0%B0%D0%BD%D0%B5%D0%BB%D1%8C%20%E2%80%9C%D0%9F%D0%B0%D1%80%D0%B0%D0%BC%D0%B5%D1%82%D1%80%D1%8B%E2%80%9D");


            AddPushButtonDataInPullDown(
                "Открыть окно переноса параметров",
                "Перенести\nпараметры",
                "Копирование значений из параметра в параметр",
                string.Format(
                    "Есть возможность сохранения и выбора файлов настроек\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                typeof(ExternalCommands.CommandCopyElemParamData).FullName,
                paramsPullDownBtn,
                "paramSetter",
                "http://moodle/mod/book/view.php?id=502&chapterid=992");


            AddPushButtonDataInPullDown(
                "Параметры\nзахваток",
                "Параметры\nзахваток",
                "Производит заполнение параметров Секции и Этажа по требованиям ВЕР под проект",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                typeof(ExternalCommands.CommandGripParam).FullName,
                paramsPullDownBtn,
                "gripParams",
                "http://moodle/mod/book/view.php?id=502&chapterid=992#:~:text=%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%9E%D0%94%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%22-,%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%97%D0%90%D0%A5%D0%92%D0%90%D0%A2%D0%9E%D0%9A%22,-%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%D0%90");
            #endregion

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

        /// <summary>
        /// Тонкая настройка видимости текста RibbonItem
        /// </summary>
        private static void SetRIShowText(RibbonItem ri, bool showName)
        {
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(ri.GetId());
            revitRibbonItem.ShowText = showName;
        }
    }
}
