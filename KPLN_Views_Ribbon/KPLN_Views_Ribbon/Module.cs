using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using KPLN_Loader.Common;
using System.Reflection;

namespace KPLN_Views_Ribbon
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class Module : IExternalModule
    {
        public static string AssemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Виды и листы");

            // Добавляю выпадающий список pullDown
            PulldownButtonData pullDownData = new PulldownButtonData("Views", "Виды");
            pullDownData.ToolTip = "Пакетная работа с видами";
            PulldownButton pullDown = panel.AddItem(pullDownData) as PulldownButton;
            BtnImagine(pullDown, "mainButton.png");

            // Добавляю в выпадающий список элементы
            AddPushButtonDataInPullDown(
                "BatchCreate",
                "Создать\nфильтры",
                "Создать фильтры",
                string.Format(
                    "Создает набор фильтров графики про критериям из CSV-файла. Примеры файлов в папке с программой.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandCreate).FullName,
                pullDown,
                "CommandCreate_small.png",
                "http://moodle.stinproject.local"
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
                typeof(ExternalCommands.CommandBatchDelete).FullName,
                pullDown,
                "CommandBatchDelete_small.png",
                "http://moodle.stinproject.local"
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
                typeof(ExternalCommands.CommandViewColoring).FullName,
                pullDown,
                "CommandViewColoring_small.png",
                "http://moodle.stinproject.local"
            );

            
            AddPushButtonDataInPullDown(
                "WallHatch",
                "Штриховки\nстен",
                "Штриховка по высоте стен",
                string.Format(
                    "Вычисляет отметки верха и низа стен; создает набор фильтров и выделяет стены разными штриховками;\n" +
                    "записывает условное обозначение, соответствующее штриховке.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.CommandWallHatch).FullName,
                pullDown,
                "CommandWallHatch_small.png",
                "http://bim-starter.com/plugins/wallhatch/"
            );

            return Result.Succeeded;
        }

        public Result Close()
        {
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
            PushButtonData data = new PushButtonData(name, text, AssemblyPath, className);
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
            string imageFullPath = Path.Combine(new FileInfo(AssemblyPath).DirectoryName, @"Resources\", imageName);
            button.LargeImage = new BitmapImage(new Uri(imageFullPath));

        }
    }
}
