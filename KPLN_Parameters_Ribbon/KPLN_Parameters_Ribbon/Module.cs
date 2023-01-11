using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using static KPLN_Loader.Output.Output;

namespace KPLN_Parameters_Ribbon
{
    public class Module : IExternalModule
    {
        private readonly string _AssemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Параметры");

            //Добавляю кнопку в панель
            AddPushButtonDataInPanel(
                "Открыть окно переноса параметров",
                "Перенести\nпараметры",
                "Копирование значений из параметра в параметр",
                string.Format(
                    "Есть возможность сохранения и выбора файлов настроек\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandCopyElemParamData).FullName,
                panel,
                "paramSetter.png",
                "http://moodle"
            );

            //Добавляю кнопку в панель
            AddPushButtonDataInPanel(
                "Копирование параметров проекта",
                "Параметры\nпроекта",
                "Производит копирование параметров проекта из файла, содержащего сведения о проекте.\n" +
                    "Для копирования параметров необходимо открыть исходный файл или подгрузить его как связь.",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandCopyProjectParams).FullName,
                panel,
                "copyProjectParams.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=992#:~:text=%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%D0%90%22-,%D0%9F%D0%A3%D0%A2%D0%AC,-%D0%9F%D0%B0%D0%BD%D0%B5%D0%BB%D1%8C%20%E2%80%9C%D0%9F%D0%B0%D1%80%D0%B0%D0%BC%D0%B5%D1%82%D1%80%D1%8B%E2%80%9D"
            );

            //Добавляю выпадающий список в панель
            PulldownButtonData pullDownData = new PulldownButtonData("Параметры под проект", "Параметры\nпод проект");
            pullDownData.ToolTip = "Коллекция плагинов, для заполнения парамтеров под конкретный проект";
            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            BtnImagine(pullDown, "paramPullDown.png");

            //Добавляю кнопку в выпадающий список pullDown
            AddPushButtonDataInPullDown(
            "Параметры захваток",
            "Параметры захваток",
            "Производит заполнение параметров Секции и Этажа по требованиям ВЕР под проект",
            string.Format(
                "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                ModuleData.Date,
                ModuleData.Version,
                ModuleData.ModuleName
            ),
            typeof(ExternalCommands.CommandGripParam).FullName,
            pullDown,
            "gripParams.png",
            "http://moodle/mod/book/view.php?id=502&chapterid=992#:~:text=%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%9E%D0%94%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%22-,%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%97%D0%90%D0%A5%D0%92%D0%90%D0%A2%D0%9E%D0%9A%22,-%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9F%D0%90%D0%A0%D0%90%D0%9C%D0%95%D0%A2%D0%A0%D0%AB%20%D0%9F%D0%A0%D0%9E%D0%95%D0%9A%D0%A2%D0%90"
        );
            

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для добавления кнопки в выпадающий список
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="pullDownButton">Выпадающий список, в который добавляем кнопку</param>
        /// <param name="imageName">Имя иконки</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPullDown(string name, string text, string shortDescription, string longDescription, string className, PulldownButton pullDownButton, string imageName, string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _AssemblyPath, className);
            PushButton button = pullDownButton.AddPushButton(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            BtnImagine(button, imageName);
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
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _AssemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            BtnImagine(button, imageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="button">Кнопка, куда нужно добавить иконку</param>
        /// <param name="imageName">Имя иконки с раширением</param>
        private void BtnImagine(RibbonButton button, string imageName)
        {
            string imageFullPath = Path.Combine(new FileInfo(_AssemblyPath).DirectoryName, @"Imagens\", imageName);
            button.LargeImage = new BitmapImage(new Uri(imageFullPath));

        }
    }
}
