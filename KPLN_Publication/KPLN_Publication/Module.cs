using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using KPLN_Loader.Common;
using Autodesk.Revit.Attributes;
using System.Reflection;

namespace KPLN_Publication
{
    [Transaction(TransactionMode.Manual)]
    public class Module : IExternalModule
    {
        public static string assemblyPath = "";

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPathname)
        {
            System.IO.Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return decoder.Frames[0];
        }

        private void AddPushButtonData(PulldownButton pullDown, string name, string text, string className, string largeImage, string image, string tTip, string lDiscr)
        {
            PushButton button = pullDown.AddPushButton(new PushButtonData(
                name, 
                text, 
                Assembly.GetExecutingAssembly().Location, 
                className)
                ) as PushButton;
            button.LargeImage = PngImageSource(largeImage);
            button.Image = PngImageSource(image);
            button.ToolTip = tTip;
            button.LongDescription = lDiscr;
        }

        public Result Execute(UIControlledApplication application, string tabName)
        {
            try { application.CreateRibbonTab(tabName); } catch { }

            assemblyPath = Assembly.GetExecutingAssembly().Location;
            RibbonPanel panel1 = application.CreateRibbonPanel(tabName, "Печать, публикация");
            PushButton btnCreate = panel1.AddItem(new PushButtonData(
                "CreateHoleTask",
                "Пакетная\nпечать",
                assemblyPath,
                "KPLN_Publication.ExternalCommands.Print.CommandBatchPrint")
                ) as PushButton;
            btnCreate.LargeImage = PngImageSource("KPLN_Publication.Resources.PrintBig.png");
            btnCreate.Image = PngImageSource("KPLN_Publication.Resources.PrintSmall.png");
            btnCreate.ToolTip = "Пакетная печать выбранных листов «на бумагу» или в PDF с автоматическим разделением по форматам.";
            btnCreate.LongDescription = "Возможности:\n" +
                " - Автоматическое определение форматов;\n" +
                " - Печать «на бумагу» и в формат PDF;\n" +
                " - Обработка нестандартных форматов - А2х3 и любых произвольных размеров (нужны права администратора);\n" +
                " - Черно/белая печать с преобразованием в черный всех цветов кроме выбранных;\n" +
                " - Печать спецификаций, разделенных на несколько листов;\n" +
                " - Печать листов из связанного файла;\n" +
                " - Объединение листов в один PDF;\n" +
                " - Авто именование PDF файлов по маске;\n\n";
            btnCreate.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=667"));
            // Stacked items: Обновить спецификации, создание наборов публикаций и набор действий перед выдачей
            // Обновить спецификации
            PushButtonData btnRefresh = new PushButtonData("RefreshSchedules", "Обновить\nспецификации", assemblyPath, "KPLN_Publication.ExternalCommands.BeforePublication.CommandRefreshSchedules");
            btnRefresh.LargeImage = PngImageSource("KPLN_Publication.Resources.UpdateBig.png");
            btnRefresh.Image = PngImageSource("KPLN_Publication.Resources.UpdateSmall.png");
            btnRefresh.ToolTip = "...";
            btnRefresh.LongDescription = "Обновляет спецификации на листах";
            // Создание наборов публикаций
            PushButtonData btnPublSets = new PushButtonData("CreatePublicationSets", "Менеджер\nнаборов", assemblyPath, "KPLN_Publication.ExternalCommands.BeforePublication.CommandOpenSetManager");
            btnPublSets.LargeImage = PngImageSource("KPLN_Publication.Resources.SetsBig.png");
            btnPublSets.Image = PngImageSource("KPLN_Publication.Resources.SetsSmall.png");
            btnPublSets.LongDescription = "Пакетно создает наборы публикации по определенным условиям";
            btnPublSets.ToolTip = "Утилита для создания наборов видов и листов (для печати и экспорта DWG/BIM360)";
            btnPublSets.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle.stinproject.local/mod/book/view.php?id=502&chapterid=666"));
            // Набор действий перед выдачей
            PulldownButtonData pullDownData = new PulldownButtonData("Перед выдачей", "Перед выдачей");
            pullDownData.LargeImage = PngImageSource("KPLN_Publication.Resources.PublBig.png");
            pullDownData.Image = PngImageSource("KPLN_Publication.Resources.PublSmall.png");
            // Stacked items: добавление элементов
            IList<RibbonItem> stackedGroup = panel1.AddStackedItems(btnRefresh, btnPublSets, pullDownData);
            PulldownButton pullDownPubl = stackedGroup[2] as PulldownButton;
            AddPushButtonData(pullDownPubl, 
                "delLists", 
                "Удалить листы",
                "KPLN_Publication.ExternalCommands.BeforePublication.CommandDelLists",
                "KPLN_Publication.Resources.DeleteLists.png",
                "KPLN_Publication.Resources.DeleteLists.png", 
                "Удаляет листы, которые не входят в параметры публикации",
                "Выбираешь параметры публикации (можно несколько), которые будешь передавать Заказчику и все листы, которые в них не входят - удалятся");
            AddPushButtonData(pullDownPubl,
                "delViews",
                "Удалить виды не на листах",
                "KPLN_Publication.ExternalCommands.BeforePublication.CommandDelViews",
                "KPLN_Publication.Resources.DeleteViews.png",
                "KPLN_Publication.Resources.DeleteViews.png",
                "Удаляет виды, которые НЕ расположены на листах",
                "Скрипт выдаёт список НЕ размещенных на листы видов (план этажа/потолка, 3d-вид, чертежный вид, спецификации). Все позиции которые выберешь - удалятся");

            return Result.Succeeded;
        }

        public Result Close()
        {
            return Result.Succeeded;
        }
    }
}
