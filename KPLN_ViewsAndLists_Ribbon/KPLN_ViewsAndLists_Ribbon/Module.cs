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

namespace KPLN_ViewsAndLists_Ribbon
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

            // Добавляю выпадающие списки pullDown для видов
            PulldownButtonData pullDownData_Views = new PulldownButtonData("Views", "Виды");
            pullDownData_Views.ToolTip = "Пакетная работа с видами";
            PulldownButton pullDown_Views = panel.AddItem(pullDownData_Views) as PulldownButton;
            BtnImagine(pullDown_Views, "mainViews.png");

            // Добавляю выпадающие списки pullDown для листов
            PulldownButtonData pullDownData_Lists = new PulldownButtonData("Lists", "Листы");
            pullDownData_Lists.ToolTip = "Пакетная работа с листами";
            PulldownButton pullDown_Lists = panel.AddItem(pullDownData_Lists) as PulldownButton;
            BtnImagine(pullDown_Lists, "mainLists.png");

            #region Добавляю в выпадающий список элементы для видов
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
                typeof(ExternalCommands.Views.CommandCreate).FullName,
                pullDown_Views,
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
                typeof(ExternalCommands.Views.CommandBatchDelete).FullName,
                pullDown_Views,
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
                typeof(ExternalCommands.Views.CommandViewColoring).FullName,
                pullDown_Views,
                "CommandViewColoring_small.png",
                "http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=671l"
            );

            
            AddPushButtonDataInPullDown(
                "WallHatch",
                "Штриховки\nстен",
                "Штриховка по высоте стен",
                string.Format(
                    "Вычисляет отметки верха и низа стен;\n"+
                    "Создает набор фильтров и выделяет стены разными штриховками;\n" +
                    "Записывает условное обозначение, соответствующее штриховке.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Views.CommandWallHatch).FullName,
                pullDown_Views,
                "CommandWallHatch_small.png",
                "http://bim-starter.com/plugins/wallhatch/"
            );
            #endregion

            #region Добавляю в выпадающий список элементы для листов
            AddPushButtonDataInPullDown(
                "RenumberLists",
                "Перенумеровать\nлисты",
                "Перенумеровать листы",
                string.Format(
                    "Изменяет нумерацию по заданной функции;\n" +
                    "Заполняет выбранный символ Юникода.\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExternalCommands.Lists.CommandListRename).FullName,
                pullDown_Lists,
                "CommandListRename.png",
                "http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=911"
            );
            #endregion

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
