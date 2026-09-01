using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Loader.Common;
using KPLN_MEPBender.ExternalCommands;
using KPLN_TrailingMEP.ExternalCommands;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KPLN_TrailingMEP
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
            // Установка основных полей модуля
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Настраиваю доступ тут, т.к. через БД много строк добавлять
            if (SQLiteMainService.CurrentUserDBSubDepartment.Id != 2 && SQLiteMainService.CurrentUserDBSubDepartment.Id != 3)
            {
                string panelName = "ИОС: Трассировка";
                RibbonPanel panel = GetOrCreatePanel(application, tabName, panelName);

                AddMepRouteStack(panel);
            }

            return Result.Succeeded;
        }

        private RibbonPanel GetOrCreatePanel(UIControlledApplication application, string tabName, string panelName)
        {
            IEnumerable<RibbonPanel> tryPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryPanels.Any())
                return tryPanels.FirstOrDefault();

            return application.CreateRibbonPanel(tabName, panelName);
        }

        private void AddMepRouteStack(RibbonPanel panel)
        {
            Assembly benderAssembly = typeof(MepBenderExtCmd).Assembly;
            string benderAssemblyName = benderAssembly.GetName().Name;
            string[] stackItemImgs = new string[] { "pipeLine", "mepBender" };
            string[] stackItemAssemblies = new string[] { _assemblyName, benderAssemblyName };

            PushButtonData trailingButtonData = CreatePushButtonData(
                string.Join("\n", RouteTraceExtCmd.PluginName.Split(' ')),
                string.Join("\n", RouteTraceExtCmd.PluginName.Split(' ')),
                "Продолжить выбранные трубы, воздуховоды и кабельные лотки до указанной точки",
                string.Format(
                    "Создает редактируемую линию траектории, а затем строит продолжения выбранного пучка с исходными типами, системами и смещениями.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(RouteTraceExtCmd).FullName,
                _assemblyPath,
                _assemblyName,
                stackItemImgs[0],
                "http://moodle.stinproject.local"
            );

            PushButtonData benderButtonData = CreatePushButtonData(
                string.Join("\n", MepBenderExtCmd.PluginName.Split(' ')),
                string.Join("\n", MepBenderExtCmd.PluginName.Split(' ')),
                "Создание обходов для труб, воздуховодов и кабельных лотков",
                string.Format(
                    "Изменяет трассировку выбранных MEP-элементов вокруг выбранных препятствий.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    benderAssemblyName
                ),
                typeof(MepBenderExtCmd).FullName,
                benderAssembly.Location,
                benderAssemblyName,
                stackItemImgs[1],
                "http://moodle.stinproject.local"
            );

            IList<RibbonItem> stackedGroup = panel.AddStackedItems(trailingButtonData, benderButtonData);
            PrepareStackedButtons(stackedGroup, stackItemImgs, stackItemAssemblies);
        }

        /// <summary>
        /// Метод для подготовки данных кнопки
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="imageName">Имя иконки. Формат имени "Имя16.png", "Имя16_dark.png"</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private PushButtonData CreatePushButtonData(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            string assemblyPath,
            string assemblyName,
            string imageName,
            string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, assemblyPath, className)
            {
                ToolTip = shortDescription,
                LongDescription = longDescription,
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(assemblyName, imageName, 16),
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(assemblyName, imageName, 32)
            };

            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            return data;
        }

        private void PrepareStackedButtons(IList<RibbonItem> stackedGroup, string[] stackItemImgs, string[] stackItemAssemblies)
        {
            for (int i = 0; i < stackedGroup.Count; i++)
            {
                RibbonItem item = stackedGroup[i];
                object parentId = typeof(RibbonItem)
                    .GetField("m_parentId", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.GetValue(item) ?? string.Empty;

                MethodInfo generateIdMethod = typeof(RibbonItem)
                    .GetMethod("generateId", BindingFlags.Static | BindingFlags.NonPublic);

                string itemId = (string)generateIdMethod?.Invoke(item, new[] { parentId, item.Name });
                if (string.IsNullOrEmpty(itemId))
                    continue;

                var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(itemId);
                if (revitRibbonItem == null)
                    continue;

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
                // Регистрация кнопки для смены иконок
                KPLN_Loader.Application.KPLNStackButtonsForImageReverse.Add((revitRibbonItem, stackItemImgs[i], stackItemAssemblies[i]));
#endif
            }
        }
    }
}
