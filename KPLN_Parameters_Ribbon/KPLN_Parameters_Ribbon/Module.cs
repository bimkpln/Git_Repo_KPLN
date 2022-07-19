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
                typeof(ExternalCommands.CommandOpenSetManager).FullName,
                panel,
                "paramSetter.png",
                "http://moodle.stinproject.local"
            );

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
                "http://moodle.stinproject.local"
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
