using Autodesk.Revit.UI;
using KPLN_ExtraFilter.Common;
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
            //Добавляю панель
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Выбор элементов");

            PushButtonData btnSelectByClick = new PushButtonData(
                "По элементу",
                "По элементу",
                _assemblyPath,
                typeof(ExternalCommands.SelectionByClickExtCommand).FullName)
            {
                LargeImage = PngImageSource("KPLN_ExtraFilter.Imagens.ClickLarge.png"),
                Image = PngImageSource("KPLN_ExtraFilter.Imagens.ClickSmall.png"),
                ToolTip = "Для выбора элементов в проекте, которые похожи/связаны с выбранным.\nВАЖНО: Сначала выдели 1 эл-т.",
                LongDescription = "Выделяешь элемент в проекте, и выбираешь сценарий, по которому будет осуществлен поиск подобных элементов" +
                    "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                AvailabilityClassName = typeof(ButtonAvailable).FullName,
            };
            btnSelectByClick.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, "http://moodle.stinproject.local"));

            PushButtonData btnSetPramsByFrame = new PushButtonData(
                "Задать рамкой",
                "Задать рамкой",
                _assemblyPath,
                typeof(ExternalCommands.SetParamsByFrameExtCommand).FullName)
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

            return Result.Succeeded;
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
