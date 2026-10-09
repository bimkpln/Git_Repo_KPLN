using Autodesk.Revit.UI;
using KPLN_CalculateTEP.ExternalCommands;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_CalculateTEP
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close() => Result.Succeeded;

        public Result Execute(UIControlledApplication application, string tabName)
        {
            // Установка основных полей модуля
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            // Ищу существующее меню инструментов АР
            if (!AddCalculateTEPButton(application.GetRibbonPanels(tabName)))
            {
                // Tools может загружаться позже. Использую общую очередь Loader.
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new AddRibbonButtonCommand(this, tabName));
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Добавление расчёта ТЭП в существующий выпадающий список АР
        /// </summary>
        private bool AddCalculateTEPButton(IList<RibbonPanel> panels)
        {
            RibbonPanel panel = panels.FirstOrDefault(i => i.Name == "Инструменты");
            PulldownButton arToolsPullDownBtn = panel?.GetItems()
                .OfType<PulldownButton>()
                .FirstOrDefault(i => i.Name == "Плагины АР");
            if (arToolsPullDownBtn == null)
                return false;

            if (arToolsPullDownBtn.GetItems().Any(i => i.Name == "Command_AR_CalculateTEP"))
                return true;

            // Revit displays LargeImage in a pulldown too; use the same 16px icon as the other menu items.
            PushButtonData calculateTEP = CreateBtnData(
                "Command_AR_CalculateTEP",
                "Расчёт ТЭП",
                "Расчёт технико-экономических показателей по старой или новой методике",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                typeof(Command_AR_CalculateTEP).FullName,
                "KPLN_CalculateTEP.Imagens.CalculateTEPSmall.png",
                "KPLN_CalculateTEP.Imagens.CalculateTEPSmall.png",
                "http://moodle");
            arToolsPullDownBtn.AddPushButton(calculateTEP);

            return true;
        }

        /// <summary>
        /// Метод для создания PushButtonData будущей кнопки
        /// </summary>
        private PushButtonData CreateBtnData(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            string smlImageName,
            string lrgImageName,
            string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className)
            {
                Text = text,
                ToolTip = shortDescription,
                LongDescription = longDescription,
            };
            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            data.Image = PngImageSource(smlImageName);
            data.LargeImage = PngImageSource(lrgImageName);

            return data;
        }

        /// <summary>
        /// Метод для добавления иконки ButtonData
        /// </summary>
        /// <param name="embeddedPathname">Имя иконки. Для иконок указать Build Action -> Embedded Resource</param>
        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        /// <summary>
        /// Добавление кнопки после завершения загрузки модулей
        /// </summary>
        private class AddRibbonButtonCommand : IExecutableCommand
        {
            private readonly Module _module;
            private readonly string _tabName;

            internal AddRibbonButtonCommand(Module module, string tabName)
            {
                _module = module;
                _tabName = tabName;
            }

            public Result Execute(UIApplication app)
            {
                return _module.AddCalculateTEPButton(app.GetRibbonPanels(_tabName))
                    ? Result.Succeeded
                    : Result.Cancelled;
            }
        }
    }
}
