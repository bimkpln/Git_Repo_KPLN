using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace KPLN_ModelChecker_Debugger
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
            //Добавляю кнопку в панель
            string currentPanelName = "Контроль качества";
            RibbonPanel currentPanel = application.GetRibbonPanels(tabName).Where(i => i.Name == currentPanelName).ToList().FirstOrDefault() ?? application.CreateRibbonPanel(tabName, "Контроль качества");

            //Добавляю выпадающий список pullDown
            PulldownButtonData pullDownData = new PulldownButtonData("Исправить", "Исправить")
            {
                ToolTip = "Набор плагинов, для исправления выявленных ошибок в модели"
            };
            PulldownButton pullDown = currentPanel.AddItem(pullDownData) as PulldownButton;
            BtnImagine(pullDown, "mainLarge.png");

            //Добавляю pinner в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "Прикрепить",
                "Прикрепить элементы модели",
                "Прикрепляет (pin) следующие элементы: связи, оси, уровни, базовую точку проекта",
                string.Format(
                    "Прикрпление необходимо для избежания случайного перемещения объектов, что может привести к проектным ошибкам.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Pinner).FullName,
                pullDown,
                "pinnerLarge.png",
                "http://moodle/mod/page/view.php?id=189"
            );

            //Добавляю worksetter в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "Рабочие наборы",
                "Рабочие наборы",
                "Распределяет элементы по рабочим наборам",
                string.Format(
                    "Возможности:\nСоздание рабочих наборов и распределение элементов по ним по настроенным правилам. Примеры файлов в папке с программой.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Worksetter).FullName,
                pullDown,
                "worksetLarge.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=668"
            );

            //Добавляю LevelAndGridsParamCopier в выпадающий список pullDown
            AddPushButtonDataInPullDown(
                "Копировать параметры сеток",
                "Копировать параметры сеток",
                "Копирует параметры и их значения для уровней и осей из разбивочного файла. ",
                string.Format(
                    "1. Возможности:\nКопирование занчения параметров 'На уровень выше';" +
                    "\n2. Копирование параметров и их значений для заполнения захваток." +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.LevelAndGridsParamCopier).FullName,
                pullDown,
                "copyProjectParams.png",
                "http://moodle/"
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
