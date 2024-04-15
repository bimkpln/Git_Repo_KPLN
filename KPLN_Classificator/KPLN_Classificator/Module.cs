using Autodesk.Revit.UI;
using KPLN_Classificator.ExecutableCommand;
using KPLN_Loader.Common;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media.Imaging;
using static KPLN_Classificator.ApplicationConfig;

namespace KPLN_Classificator
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Module : IExternalModule
    {
        public static string assemblyPath = "";

        public Module()
        {
            //Конфигурирование приложения для работы в среде KPLN_Loader
            output = new KplnOutput();
            commandEnvironment = new KplnCommandEnvironment();
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            try { application.CreateRibbonTab(tabName); } catch { }

            string panelName = "Параметры";
            RibbonPanel panel = null;
            List<RibbonPanel> tryPanels = application.GetRibbonPanels(tabName).Where(i => i.Name == panelName).ToList();
            if (tryPanels.Count == 0)
            {
                panel = application.CreateRibbonPanel(tabName, panelName);
            }
            else
            {
                panel = tryPanels.First();
            }

            PushButtonData btnHostMark = new PushButtonData(
                "ClassificatorCompleteCommand",
                "Заполнить\nклассификатор",
                assemblyPath,
                typeof(CommandOpenClassificatorForm).FullName);

            btnHostMark.LargeImage = PngImageSource("KPLN_Classificator.Resources.Classificator_large.PNG");
            btnHostMark.Image = PngImageSource("KPLN_Classificator.Resources.Classificator.PNG");
            btnHostMark.ToolTip = "Параметризация элементов согласно заданным правилам.";
            btnHostMark.LongDescription = "Возможности:\n" +
                "Задание правил для параметризации элементов;\n" +
                "Маппинг параметров (передача значений между параметрами элемента);\n" +
                "Сохранение конфигурационного файла с возможностью повторного использования;\n";
            btnHostMark.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle/mod/book/view.php?id=502&chapterid=669"));
            btnHostMark.AvailabilityClassName = typeof(Availability.StaticAvailable).FullName;

            panel.AddItem(btnHostMark);

            return Result.Succeeded;
        }

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPathname)
        {
            System.IO.Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);

            PngBitmapDecoder decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }

        public Result Close()
        {
            return Result.Succeeded;
        }
    }
}
