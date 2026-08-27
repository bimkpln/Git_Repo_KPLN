using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Loader.Common;
using KPLN_MEPBender.ExternalCommands;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KPLN_MEPBender
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;
        private readonly string _assemblyName = Assembly.GetExecutingAssembly().GetName().Name;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Настраиваю доступ тут, т.к. через БД много строк добавлять
            if (SQLiteMainService.CurrentUserDBSubDepartment.Id != 2 && SQLiteMainService.CurrentUserDBSubDepartment.Id != 3)
            {
                //Ищу или создаю панель инструменты
                string panelName = "Междисциплинарный анализ";
                RibbonPanel panel = null;
                IEnumerable<RibbonPanel> tryPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
                if (tryPanels.Any())
                    panel = tryPanels.FirstOrDefault();
                else
                    panel = application.CreateRibbonPanel(tabName, panelName);


                AddPushButtonDataInPanel(
                    string.Join("\n", MepBenderExtCmd.PluginName.Split(' ')),
                    string.Join("\n", MepBenderExtCmd.PluginName.Split(' ')),
                    "Создание обходов для труб, воздуховодов и кабельных лотков",
                    string.Format(
                        "Изменяет трассировку выбранных MEP-элементов вокруг выбранных препятствий.\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName),
                    typeof(MepBenderExtCmd).FullName,
                    panel,
                    "mepBender",
                    "http://moodle.stinproject.local"
                );
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для добавления отдельной кнопки в панель
        /// </summary>
        private void AddPushButtonDataInPanel(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            RibbonPanel panel,
            string imageName,
            string contextualHelp)
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
