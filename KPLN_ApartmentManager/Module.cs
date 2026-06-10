using Autodesk.Revit.UI;
using KPLN_ApartmentManager.ExecutableCommand;
using KPLN_Library_DBWorker;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_ApartmentManager
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

            //Ищу или создаю панель
            string panelName = "Менеджеры";
            RibbonPanel mPanel = null;
            IEnumerable<RibbonPanel> tryMPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName);
            if (tryMPanels.Any())
                mPanel = tryMPanels.FirstOrDefault();
            else
                mPanel = application.CreateRibbonPanel(tabName, panelName);

            if (SQLiteMainService.CurrentUserDBSubDepartment.Id == 2 || SQLiteMainService.CurrentUserDBSubDepartment.Id == 8)
            {
                PushButtonData apartmentManager = CreateBtnData(
                ExecutableCommand.CommandApartmentManagerShow.PluginName,
                ExecutableCommand.CommandApartmentManagerShow.PluginName,
                "Менеджер квартир",
                string.Format(
                    "Менеджер квартир.\n" +
                    "\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(CommandApartmentManagerShow).FullName,
                "KPLN_ApartmentManager.Imagens.apartmentManagerBig.png",
                "KPLN_ApartmentManager.Imagens.apartmentManagerBig.png",
                "http://moodle/");
                mPanel.AddItem(apartmentManager);
            }


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
                data.AvailabilityClassName = typeof(Common.StaticAvailable).FullName;
            }

            return data;
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
