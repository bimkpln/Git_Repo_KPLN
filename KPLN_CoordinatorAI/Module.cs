using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_CoordinatorAI.Common;
using KPLN_CoordinatorAI.ExternalCommands;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_CoordinatorAI
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
            // Установка основных полей модуля
            ModuleData.RevitMainWindowHandle = application.MainWindowHandle;
            ModuleData.RevitVersion = int.Parse(application.ControlledApplication.VersionNumber);

            //Добавляю панель
            string panelName = "Инструменты";
            RibbonPanel mPanel = null;
            IEnumerable<RibbonPanel> tryMPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryMPanels.Any())
                mPanel = tryMPanels.FirstOrDefault();
            else
                mPanel = application.CreateRibbonPanel(tabName, panelName);

            PushButtonData coordinatorAI = CreateBtnData(
                    CommandCoordinatorAI.PluginName,
                    CommandCoordinatorAI.PluginName,
                    "AI-помошник при работе в Revit",
                    string.Format(
                        "AI-помошник при работе в Revit.\n" +
                        "\n" +
                        "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(CommandCoordinatorAI).FullName,
                    "KPLN_CoordinatorAI.Imagens.coordinatorAISmall.png",
                    "KPLN_CoordinatorAI.Imagens.coordinatorAIBig.png",
                    "http://moodle/mod/book/view.php?id=502&chapterid=1342");

            mPanel.AddItem(coordinatorAI);

            return Result.Succeeded;
        }

        /// <summary>
        /// Метод для создания PushButtonData будущей кнопки
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private PushButtonData CreateBtnData(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string className,
            string smlImageName,
            string lrgImageName,
            string contextualHelp,
            bool avclass = false)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className)
            {
                Text = text,
                ToolTip = shortDescription,
                LongDescription = longDescription
            };
            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            data.Image = PngImageSource(smlImageName);
            data.LargeImage = PngImageSource(lrgImageName);
            if (avclass)
            {
                data.AvailabilityClassName = typeof(StaticAvailable).FullName;
            }

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
    }
}