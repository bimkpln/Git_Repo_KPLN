using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Reflection;

namespace KPLN_ExtraFilter
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Выбор элементов");

            // Отдельная кнопка
            AddPushButtonDataInPanel(
                string.Join("\n", SelectionByModelExtCmd.PluginName.Split(' ')),
                string.Join("\n", SelectionByModelExtCmd.PluginName.Split(' ')),
                "Генерация древовидной структуры из элементов модели",
                string.Format("\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(SelectionByModelExtCmd).FullName,
                panel,
                "treeModel",
                "treeModel" +
                "",
                "http://moodle/mod/book/view.php?id=502&chapterid=1341"
            );


            // Кнопки в стэк
            // Коллекция изображений, чтобы сделать замену при смене фона ревит.
            // ВАЖЕН ПОРЯДОК - он должен совпадать с порядком добавления кнопок с стэк
            string[] stackItemImgs = new string[] { "click", "frame" };

            PushButtonData btnSelectByClick = new PushButtonData(
                SelectionByClickExtCmd.PluginName,
                SelectionByClickExtCmd.PluginName,
                _assemblyPath,
                typeof(SelectionByClickExtCmd).FullName)
            {
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, stackItemImgs[0], 32),
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, stackItemImgs[0], 16),
                ToolTip = "Для выбора элементов в проекте, которые похожи/связаны с выбранным.",
                LongDescription = string.Format("Выделяешь элемент в проекте, и выбираешь сценарий, по которому будет осуществлен поиск подобных элементов" +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
#if Debug2020 || Revit2020
                AvailabilityClassName = typeof(ButtonAvailable).FullName,
#endif
            };
            btnSelectByClick.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://moodle/mod/book/view.php?id=502&chapterid=1341"));


            PushButtonData btnSetPramsByFrame = new PushButtonData(
                SetParamsByFrameExtCmd.PluginName,
                SetParamsByFrameExtCmd.PluginName,
                _assemblyPath,
                typeof(SetParamsByFrameExtCmd).FullName)
            {
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, stackItemImgs[1], 32),
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, stackItemImgs[1], 16),
                ToolTip = "Добавляет к выбранным элементам ВСЕ вложенности, а также может заполнить параметры для ВСЕХ вложенностей",
                LongDescription = string.Format("При выделении дополнительно выделятся:" +
                    "\n  1. Вложенные элементы семейств." +
                    "\n  2. Изоляция воздуховодов/труб." +
                    "\n  3. Отдельные элементы групп (важно, чтобы параметр мог меняться по экземплярам групп)." +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
            };
            btnSetPramsByFrame.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://moodle/mod/book/view.php?id=502&chapterid=1341"));


            IList<RibbonItem> stackedGroup = panel.AddStackedItems(btnSelectByClick, btnSetPramsByFrame);
            // Скрываю текстовое название кнопок
            for (int i = 0; i < stackedGroup.Count; i++)
            {
                RibbonItem item = stackedGroup[i];

                var parentId = typeof(RibbonItem)
                    .GetField("m_parentId", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.GetValue(item) ?? string.Empty;
                var generateIdMethod = typeof(RibbonItem)
                    .GetMethod("generateId", BindingFlags.Static | BindingFlags.NonPublic);

                string itemId = (string)generateIdMethod?.Invoke(item, new[] { parentId, item.Name });

                var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(itemId);
                revitRibbonItem.ShowText = false;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
                // Регистрация кнопки для смены иконок
                KPLN_Loader.Application.KPLNStackButtonsForImageReverse.Add((revitRibbonItem, stackItemImgs[i], Assembly.GetExecutingAssembly().GetName().Name));
#endif
            }

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
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string bigImageName, string contextualHelp)
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
