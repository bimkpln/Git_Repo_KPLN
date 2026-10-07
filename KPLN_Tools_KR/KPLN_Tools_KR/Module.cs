using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools_KR.Common;
using KPLN_Tools_KR.ExternalCommands;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools_KR
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

            PulldownButton krToolsPullDownBtn = CreatePulldownButtonInRibbon(
                "Плагины КР",
                "Плагины КР",
                "КР: Коллекция плагинов для автоматизации задач",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                "krMain",
                panel,
                false);

            PushButtonData smnx_Rebar = CreateBtnData(
                "SMNX_Металоёмкость",
                "SMNX_Металоёмкость",
                "SMNX: Заполняет параметр \"SMNX_Расход арматуры (Кг/м3)\"",
                string.Format(
                    "Варианты запуска:\n" +
                        "1. Записать объём бетона и основную марку в арматуру;\n" +
                        "2. Перенести значения из спецификации в параметр;\n" +
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_KR_SMNX_RebarHelper).FullName,
                "KPLN_Tools_KR.Imagens.wipeSmall.png",
                "KPLN_Tools_KR.Imagens.wipeSmall.png",
                "http://moodle");

            PushButtonData kr_IFCRebarMark = CreateBtnData(
                ExtCmd_KR_IFCRebarMark.PluginName,
                ExtCmd_KR_IFCRebarMark.PluginName,
                "Автоматически заполняет IFC-арматуре значение параметра Мрк.МаркаКонструкции",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_KR_IFCRebarMark).FullName,
                "KPLN_Tools_KR.Imagens.IFCRebarMarkSmall.png",
                "KPLN_Tools_KR.Imagens.IFCRebarMarkSmall.png",
                "http://moodle");


#if Debug2024 || Revit2024
            PushButtonData kr_expitVolume = CreateBtnData(
                ExtCmd_KR_ExpitVolume.PluginName,
                ExtCmd_KR_ExpitVolume.PluginName,
                "Плагин для получения объема котлована",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName
                ),
                typeof(ExtCmd_KR_ExpitVolume).FullName,
                "KPLN_Tools_KR.Imagens.expitVolumeSmall.png",
                "KPLN_Tools_KR.Imagens.expitVolumeSmall.png",
                "http://moodle");

            krToolsPullDownBtn.AddPushButton(kr_expitVolume);
#endif

            krToolsPullDownBtn.AddPushButton(smnx_Rebar);
            krToolsPullDownBtn.AddPushButton(kr_IFCRebarMark);

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