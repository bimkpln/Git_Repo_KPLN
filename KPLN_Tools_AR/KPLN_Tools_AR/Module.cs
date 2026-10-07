using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools_AR.Common;
using KPLN_Tools_AR.ExternalCommands;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools_AR
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

            PulldownButton arToolsPullDownBtn = CreatePulldownButtonInRibbon(
                "Плагины АР",
                "Плагины АР",
                "АР: Коллекция плагинов для автоматизации задач",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                "arMain",
                panel,
                false);

            PushButtonData arGNSArea = CreateBtnData(
                "Площадь ГНС",
                "Площадь ГНС",
                "Обводит внешние границы здания на плане",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_AR_GNSBound).FullName,
                "KPLN_Tools_AR.Imagens.gnsAreaBig.png",
                "KPLN_Tools_AR.Imagens.gnsAreaSmall.png",
                "http://moodle");

            PushButtonData arPyatnGraph = CreateBtnData(
                "Пятнография: Экспликация",
                "Пятнография: Экспликация",
                "Проверяет помещения/цветовые облости на соответсвие ТЗ и позволяет сформировать итоговую спецификацю",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_AR_PyatnGraph).FullName,
                "KPLN_Tools_AR.Imagens.arPyatnGraphBig.png",
                "KPLN_Tools_AR.Imagens.arPyatnGraphSmall.png",
                "http://moodle");

            PushButtonData TEPDesign = CreateBtnData(
                "Оформление ТЭП",
                "Оформление ТЭП",
                "Плагин для оформления ТЭП",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_AR_TEPDesign).FullName,
                "KPLN_Tools_AR.Imagens.TEPDesignBig.png",
                "KPLN_Tools_AR.Imagens.TEPDesignSmall.png",
                "http://moodle");

            PushButtonData evacuationRoutes = CreateBtnData(
                "Автомоделирование путей эвакуации",
                "Автомоделирование путей эвакуации",
                "Плагин для автомоделирования путей эвакуации",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_AR_EvacuationRoutes).FullName,
                "KPLN_Tools_AR.Imagens.evacuationRoutesBig.png",
                "KPLN_Tools_AR.Imagens.evacuationRoutesSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1350");

            PushButtonData Furniture3DFrom2D = CreateBtnData(
                "Мебель 2D <-> 3D",
                "Мебель 2D <-> 3D",
                "Преобразование мебели из 2D в 3D и наоборот",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_AR_Furniture3DFrom2D).FullName,
                "KPLN_Tools_AR.Imagens.furniture2DFrom3DBig.png",
                "KPLN_Tools_AR.Imagens.furniture2DFrom3DSmall.png",
                "http://moodle/mod/book/view.php?id=502&chapterid=1350");


            arToolsPullDownBtn.AddPushButton(arGNSArea);
            arToolsPullDownBtn.AddPushButton(Furniture3DFrom2D);
#if Debug2023 || Revit2023
            arToolsPullDownBtn.AddPushButton(arPyatnGraph);
            arToolsPullDownBtn.AddPushButton(TEPDesign);
#endif
            arToolsPullDownBtn.AddPushButton(evacuationRoutes);

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