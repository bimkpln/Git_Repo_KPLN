using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Publication.ExternalCommands.Print;
using System.Reflection;

namespace KPLN_Publication
{
    [Transaction(TransactionMode.Manual)]
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Экспорт");


            AddPushButtonDataInPanel(
                CommandBatchPrint.PluginName,
                string.Join("\n", CommandBatchPrint.PluginName.Split(' ')),
                "Пакетный перевод выбранных листов в PDF (с автоматическим разделением по форматам), DWG",
                string.Format(
                    "Возможности:\n" +
                    " - Автоматическое определение форматов;\n" +
                    " - Печать «на бумагу» и в формат PDF;\n" +
                    " - Экспорт в DWG;\n" +
                    " - Обработка нестандартных форматов - А2х3 и любых произвольных размеров (нужны права администратора);\n" +
                    " - Черно/белая печать с преобразованием в черный всех цветов кроме выбранных;\n" +
                    " - Печать спецификаций, разделенных на несколько листов;\n" +
                    " - Печать листов из связанного файла;\n" +
                    " - Объединение листов в один PDF;\n" +
                    " - Авто именование PDF файлов по маске." +
                    "\n\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandBatchPrint).FullName,
                panel,
                "printer",
                "http://moodle/mod/book/view.php?id=502&chapterid=667",
                true
            );

            return Result.Succeeded;
        }

        public Result Close()
        {
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
        private void AddPushButtonDataInPanel(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            RibbonPanel panel,
            string imageName,
            string contextualHelp,
            bool avclass)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
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
