#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using System.Windows.Media.Imaging;


namespace KPLN_RevitWorksets
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Module : IExternalModule
    {
        public static string assemblyPath = "";

        private System.Windows.Media.ImageSource PngImageSource(string embeddedPathname)
        {
            System.IO.Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return decoder.Frames[0];
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

            PushButton btnHostMark = panel.AddItem(new PushButtonData(
                "RevitWorksetsCommand",
                "Рабочие\nнаборы",
                assemblyPath,
                "RevitWorksets.Command")
                ) as PushButton;
            btnHostMark.LargeImage = PngImageSource("RevitWorksets.Resources.Command_large.png");
            btnHostMark.Image = PngImageSource("RevitWorksets.Resources.Command_small.png");
            btnHostMark.ToolTip = "Рабочие наборы";
            btnHostMark.LongDescription = "Возможности:\n" +
                "Создание рабочих наборов и распределение элементов по ним по настроенным правилам. Примеры файлов в папке с программой;\n\n";
            btnHostMark.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle.stinproject.local/mod/book/view.php?id=396&chapterid=437"));

            return Result.Succeeded;
        }

        public Result Close()
        {
            return Result.Succeeded;
        }
    }
}
