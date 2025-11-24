using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
using KPLN_ExtraFilter.ExternalCommands;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ExtraFilter
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }

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
                "KPLN_ExtraFilter.Imagens.TreeModelSmall.png",
                "KPLN_ExtraFilter.Imagens.TreeModelLarge.png",
                "http://moodle.stinproject.local"
            );


            // Кнопки в стэк
            PushButtonData btnSelectByClick = new PushButtonData(
                SelectionByClickExtCmd.PluginName,
                SelectionByClickExtCmd.PluginName,
                _assemblyPath,
                typeof(SelectionByClickExtCmd).FullName)
            {
                LargeImage = PngImageSource("KPLN_ExtraFilter.Imagens.ClickLarge.png"),
                Image = PngImageSource("KPLN_ExtraFilter.Imagens.ClickSmall.png"),
                ToolTip = "Для выбора элементов в проекте, которые похожи/связаны с выбранным.",
                LongDescription = "Выделяешь элемент в проекте, и выбираешь сценарий, по которому будет осуществлен поиск подобных элементов" +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
#if Debug2020 || Revit2020
                AvailabilityClassName = typeof(ButtonAvailable).FullName,
#endif
            };
            btnSelectByClick.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://moodle.stinproject.local"));


            PushButtonData btnSetPramsByFrame = new PushButtonData(
                SetParamsByFrameExtCmd.PluginName,
                SetParamsByFrameExtCmd.PluginName,
                _assemblyPath,
                typeof(SetParamsByFrameExtCmd).FullName)
            {
                LargeImage = PngImageSource("KPLN_ExtraFilter.Imagens.FrameLarge.png"),
                Image = PngImageSource("KPLN_ExtraFilter.Imagens.FrameSmall.png"),
                ToolTip = "Позволяет выбрать элементы рамкой с расширенным функционалом и задать параметры",
                LongDescription = "Можно добавлять несколько параметров. При выделении дополнительно выделятся:" +
                    "\n  1. Вложенные элементы семейств." +
                    "\n  2. Отдельные элементы групп (важно, чтобы параметр мог меняться по экземплярам групп." +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
            };
            btnSetPramsByFrame.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://moodle.stinproject.local"));

            IList<RibbonItem> stackedGroup = panel.AddStackedItems(btnSelectByClick, btnSetPramsByFrame);
            // Скрываю текстовое название кнопок
            foreach (RibbonItem item in stackedGroup)
            {
                var parentId = typeof(RibbonItem)
                    .GetField("m_parentId", BindingFlags.Instance | BindingFlags.NonPublic)
                    ?.GetValue(item) ?? string.Empty;
                var generateIdMethod = typeof(RibbonItem)
                    .GetMethod("generateId", BindingFlags.Static | BindingFlags.NonPublic);

                string itemId = (string)generateIdMethod?.Invoke(item, new[] { parentId, item.Name });

                var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(itemId);
                revitRibbonItem.ShowText = false;
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
            button.Image = PngImageSource(imageName);
            button.LargeImage = PngImageSource(bigImageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="embeddedPathname">Имя иконки с раширением</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
