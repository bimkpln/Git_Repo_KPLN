using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_ModelChecker_Debugger.ExternalCommands;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ModelChecker_Debugger
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
            //Добавляю кнопку в панель
            string currentPanelName = "Контроль качества";
            RibbonPanel currentPanel = application.GetRibbonPanels(tabName).Where(i => i.Name == currentPanelName).ToList().FirstOrDefault() ?? application.CreateRibbonPanel(tabName, "Контроль качества");

            //Добавляю выпадающий список pullDown
            PulldownButtonData pullDownData = new PulldownButtonData("Исправить", "Исправить")
            {
                ToolTip = "Набор плагинов, для исправления выявленных ошибок в модели",
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, "main", 32),
            };
            PulldownButton pullDown = currentPanel.AddItem(pullDownData) as PulldownButton;


            AddPushButtonDataInPullDown(
                Pinner.PluginName,
                Pinner.PluginName,
                "Прикрепляет (pin) следующие элементы: связи, оси, уровни, базовую точку проекта",
                string.Format(
                    "Прикрпление необходимо для избежания случайного перемещения объектов, что может привести к проектным ошибкам.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(Pinner).FullName,
                pullDown,
                "pinner",
                "http://moodle/mod/page/view.php?id=189"
            );


            AddPushButtonDataInPullDown(
                WorksetCreate.PluginName,
                WorksetCreate.PluginName,
                "Создаёт и распределяет элементы по рабочим наборам",
                string.Format(
                    "Возможности:\nСоздание рабочих наборов и распределение элементов по ним по настроенным правилам. Примеры файлов в папке с программой.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(WorksetCreate).FullName,
                pullDown,
                "worksetCreate",
                "http://moodle/mod/book/view.php?id=502&chapterid=668"
            );

            // ВОЗМОЖНО УДАЛЕНИЕ РН ТОЛЬКО НАЧИНАЯ С РЕВИТ2023, до этого - нет API
#if !Debug2020 && !Revit2020
            AddPushButtonDataInPullDown(
                WorksetDelete.PluginName,
                WorksetDelete.PluginName,
                "Удаляет рабочие наборы",
                string.Format(
                    "Возможности:\nУдаление рабочих наборов.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(WorksetDelete).FullName,
                pullDown,
                "worksetDelete",
                "http://moodle/"
            );
#endif


            AddPushButtonDataInPullDown(
                PatitionFileSetter.PluginName,
                PatitionFileSetter.PluginName,
                "Проверка корректности относительно уровней (по оси Z) и привязка отдельных блоков к уровням проекта (по отметкам)",
                string.Format(
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(PatitionFileSetter).FullName,
                pullDown,
                "setPatitionalFile",
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
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = pullDownButton.AddPushButton(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            button.LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            // Регистрация кнопки для смены иконок
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((button, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
        }
    }
}
