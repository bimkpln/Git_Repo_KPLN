using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Tools_OVVK.Common;
using KPLN_Tools_OVVK.ExternalCommands;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_Tools_OVVK
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

            PulldownButton ovvkToolsPullDownBtn = CreatePulldownButtonInRibbon(
                "Плагины ОВВК",
                "Плагины ОВВК",
                "ОВВК: Коллекция плагинов для автоматизации задач",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                    ModuleData.Date,
                    ModuleData.Version,
                    ModuleData.ModuleName),
                "ovvkMain",
                panel,
                false);

            PushButtonData ovvkPipeThickness = CreateBtnData(
                ExtCmd_OVVK_PipeThickness.PluginName,
                ExtCmd_OVVK_PipeThickness.PluginName,
                "Заполнение толщины стенок труб по сортаменту",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_OVVK_PipeThickness).FullName,
                "KPLN_Tools_OVVK.Imagens.pipeThicknessSmall.png",
                "KPLN_Tools_OVVK.Imagens.pipeThicknessSmall.png",
                "http://moodle");

            PushButtonData ovvkSystemManager = CreateBtnData(
                ExtCmd_OVVK_SystemManager.PluginName,
                ExtCmd_OVVK_SystemManager.PluginName,
                "Менеджер систем ОВВК",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_OVVK_SystemManager).FullName,
                "KPLN_Tools_OVVK.Imagens.systemMangerSmall.png",
                "KPLN_Tools_OVVK.Imagens.systemMangerSmall.png",
                "http://moodle");

            PushButtonData ventilationSettingsConfigurator = CreateBtnData(
                ExtCmd_OV_VentConfigurator.PluginName,
                ExtCmd_OV_VentConfigurator.PluginName,
                "Плагин-конфигуратор семейств вент установок",
                string.Format(
                    "Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName
                ),
                typeof(ExtCmd_OV_VentConfigurator).FullName,
                "KPLN_Tools_OVVK.Imagens.ventilationSettingsConfiguratorSmall.png",
                "KPLN_Tools_OVVK.Imagens.ventilationSettingsConfiguratorSmall.png",
                "http://moodle");

            PushButtonData ovDuctThickness = CreateBtnData(
                ExtCmd_OV_DuctThickness.PluginName,
                ExtCmd_OV_DuctThickness.PluginName,
                "Заполнение толщины стенок воздуховодов по СП",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_OV_DuctThickness).FullName,
                "KPLN_Tools_OVVK.Imagens.ductThicknessSmall.png",
                "KPLN_Tools_OVVK.Imagens.ductThicknessSmall.png",
                "http://moodle");

            PushButtonData ovOzkDuctAccessory = CreateBtnData(
                ExtCmd_OV_OZKDuctAccessory.PluginName,
                ExtCmd_OV_OZKDuctAccessory.PluginName,
                "Заполнение марок клапанов ОЗК",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_OV_OZKDuctAccessory).FullName,
                "KPLN_Tools_OVVK.Imagens.ozkDuctAccessorySmall.png",
                "KPLN_Tools_OVVK.Imagens.ozkDuctAccessorySmall.png",
                "http://moodle");

            #if Revit2020 || Debug2020
            PushButtonData setInsulationPipes = CreateBtnData(
                "СЕТ: Изоляция труб",
                "СЕТ: Изоляция труб",
                "Заполнение параметров изоляции труб",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_SET_InsulationPipes).FullName,
                "KPLN_Tools_OVVK.Imagens.FillInParamSmall.png",
                "KPLN_Tools_OVVK.Imagens.FillInParamSmall.png",
                "http://moodle");
            ovvkToolsPullDownBtn.AddPushButton(setInsulationPipes);
            #endif

            PushButtonData ovvkAutonumber = CreateBtnData(
                ExtCmd_OVVK_ScheduleIncrementor.PluginName,
                ExtCmd_OVVK_ScheduleIncrementor.PluginName,
                "Нумерация спецификации ОВВК",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_OVVK_ScheduleIncrementor).FullName,
                "KPLN_Tools_OVVK.Imagens.autonumberSmall.png",
                "KPLN_Tools_OVVK.Imagens.autonumberSmall.png",
                "http://moodle");

            PushButtonData auptTagPlacer = CreateBtnData(
                ExtCmd_AUPT_TagPlacer.PluginName,
                ExtCmd_AUPT_TagPlacer.PluginName,
                "Автоматическая маркировка труб АУПТ",
                string.Format("Дата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}", ModuleData.Date, ModuleData.Version, ModuleData.ModuleName),
                typeof(ExtCmd_AUPT_TagPlacer).FullName,
                "KPLN_Tools_OVVK.Imagens.auptTagSmall.png",
                "KPLN_Tools_OVVK.Imagens.auptTagSmall.png",
                "http://moodle");

            ovvkToolsPullDownBtn.AddPushButton(ovvkPipeThickness);
            ovvkToolsPullDownBtn.AddPushButton(ovDuctThickness);
            ovvkToolsPullDownBtn.AddPushButton(ovOzkDuctAccessory);
            ovvkToolsPullDownBtn.AddPushButton(ovvkSystemManager);
            ovvkToolsPullDownBtn.AddPushButton(ventilationSettingsConfigurator);
            ovvkToolsPullDownBtn.AddPushButton(ovvkAutonumber);
            ovvkToolsPullDownBtn.AddPushButton(auptTagPlacer);

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