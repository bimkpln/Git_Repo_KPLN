using Autodesk.Revit.UI;
using KPLN_Library_DBWorker;
using KPLN_Loader.Common;
using KPLN_Tools_EOM.Common;
using KPLN_Tools_EOM.ExternalCommands;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools_EOM
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

            //Ищу или создаю панель
            const string panelName = "Инструменты";
            RibbonPanel panel = application.GetRibbonPanels(tabName).FirstOrDefault(i => i.Name == panelName)
                ?? application.CreateRibbonPanel(tabName, panelName);


#if Revit2020 || Debug2020
            PulldownButton eomToolsPullDownBtn = CreatePulldownButtonInRibbon(
                "Плагины ЭОМ",
                "Плагины ЭОМ",
                "ЭОМ: Коллекция плагинов для автоматизации задач",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                "eomMain",
                panel,
                false);

            // Импорт глобальных параметров доступен только специалистам BIM-отдела.
            if (SQLiteMainService.CurrentUserDBSubDepartment?.Id == 8)
            {
                PushButtonData globalParameters = CreateBtnData(
                    "СЕТ: Глобальные параметры",
                    "СЕТ: Глобальные параметры",
                    "Создать или обновить глобальные параметры проекта из Excel.",
                    "Лист «Глобальные параметры»: A — имя, B — число, C — формула Revit.",
                    typeof(KPLN_Tools_EOM.Common.ApartmentWiring.GlobalCommand).FullName,
                    "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                    "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                    "http://moodle");

                eomToolsPullDownBtn.AddPushButton(globalParameters);
            }


            eomToolsPullDownBtn.AddPushButton(CreateBtnData(
                "СЕТ: Немоделируемые элементы",
                "СЕТ: Немоделируемые элементы",
                "Кабели, трубы, крепления и параметры секции/этажа для проекта «Сетунь».",
                "Имя файла проекта должно начинаться с «СЕТ_1».",
                typeof(KPLN_Tools_EOM.Common.ApartmentWiring.NonModelCommand).FullName,
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "http://moodle"));
            

            eomToolsPullDownBtn.AddPushButton(CreateBtnData(
                "СЕТ: Тип отделки",
                "СЕТ: Тип отделки",
                "Заполнение типа отделки и связи с глобальными параметрами проекта «Сетунь».",
                "Имя файла проекта должно начинаться с «СЕТ_1».",
                typeof(KPLN_Tools_EOM.Common.ApartmentWiring.FinishCommand).FullName,
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "http://moodle"));
            

            eomToolsPullDownBtn.AddPushButton(CreateBtnData(
                "Сортировка спецификации",
                "Сортировка спецификации",
                "Заполнение ключа сортировки спецификации проекта «Сетунь».",
                "Имя файла проекта должно начинаться с «СЕТ_1».",
                typeof(KPLN_Tools_EOM.Common.ApartmentWiring.SortCommand).FullName,
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "http://moodle"));


            eomToolsPullDownBtn.AddPushButton(CreateBtnData(
                "СЕТ: Заполнить параметры",
                "СЕТ: Заполнить параметры",
                "СЕТ: Заполнить параметры",
                string.Format("Плагин заполняет параметр для формирования спецификации для: \n" +
                    "1. Кабельных лотков;\n" +
                    "2. Соед. деталей кабельных лотков;\n" +
                    "3. Воздуховодов (огнезащита).\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_SET_EOMParams).FullName,
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "KPLN_Tools_EOM.Imagens.FillInParamSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1319"));
#endif

            return Result.Succeeded;
        }

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
                Image = PngImageSource(smlImageName),
                LargeImage = PngImageSource(lrgImageName),
            };
            data.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            return data;
        }

        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return decoder.Frames[0];
        }

        private PulldownButton CreatePulldownButtonInRibbon(
            string name,
            string text,
            string shortDescription,
            string longDescription,
            string imageName,
            RibbonPanel panel,
            bool showName)
        {
            PulldownButton pullDownRI = panel.AddItem(new PulldownButtonData(name, text)
            {
                ToolTip = shortDescription,
                LongDescription = longDescription,
                Image = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 16),
                LargeImage = KPLN_Loader.Application.GetBtnImage_ByTheme(_assemblyName, imageName, 32),
            }) as PulldownButton;

            SetRIShowText(pullDownRI, showName);

#if !Debug2020 && !Revit2020 && !Debug2023 && !Revit2023
            KPLN_Loader.Application.KPLNButtonsForImageReverse.Add((pullDownRI, imageName, Assembly.GetExecutingAssembly().GetName().Name));
#endif
            return pullDownRI;
        }

        private static void SetRIShowText(RibbonItem ri, bool showName)
        {
            var revitRibbonItem = UIFramework.RevitRibbonControl.RibbonControl.findRibbonItemById(ri.GetId());
            revitRibbonItem.ShowText = showName;
        }
    }
}