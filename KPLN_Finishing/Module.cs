using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

using KPLN_Loader.Common;
using KPLN_Loader.Output;
using KPLN_Finishing.CommandTools;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;



namespace KPLN_Finishing
{
    public class Module : IExternalModule
    {
        public Result Close()
        {
            return Result.Succeeded;
        }
        public Result Execute(UIControlledApplication application, string tabName)
        {
            Names.assembly = Assembly.GetExecutingAssembly().FullName;
            Names.assembly_Path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString();
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Отделка");
            PushButtonData buttonData_A = new PushButtonData("Рассчитать элементы отделки", "Рассчитать элементы отделки", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.Run).FullName);
            PushButton button_A = panel.AddItem(buttonData_A) as PushButton;
            button_A.ToolTip = "Определение помещений для каждого элемента отделки в проекте";
            button_A.LongDescription = "В элементах будут заполнены параметры : «О_Id помещения», «О_Номер помещений», «О_Имя помещения»";
            button_A.ItemText = "Рассчитать\nэлементы";
            button_A.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle/mod/book/view.php?id=502&chapterid=664#:~:text=%D0%B8%D0%BD%D1%81%D1%82%D1%80%D1%83%D0%BC%D0%B5%D0%BD%D1%82%D1%8B%20%D0%B0%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D0%B8%20R%D0%B5vit-,%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9E%D0%A2%D0%94%D0%95%D0%9B%D0%9A%D0%90%22,-%D0%A1%D0%BA%D1%80%D0%B8%D0%BF%D1%82%20%D0%BF%D0%BE%20%D1%80%D0%B0%D1%81%D1%87%D0%B5%D1%82%D1%83"));
            SetIcon(button_A, "link_elements_large");
            BitmapImage toolTip0 = new BitmapImage(new Uri(string.Format(@"{0}\Source\{1}.png", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "tool_tip_00")));
            button_A.ToolTipImage = toolTip0;

            PushButtonData buttonData_B = new PushButtonData("Вручную", "Привязать элемент", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.Copy).FullName);
            SetIcon(buttonData_B, "attach_large");
            PushButtonData buttonData_C = new PushButtonData("Параметры", "Загрузить параметры", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.LoadParameters).FullName);
            SetIcon(buttonData_C, "parameters_large");
            PushButtonData buttonData_D = new PushButtonData("По цветам", "Окрасить по помещению", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.SetColor).FullName);
            SetIcon(buttonData_D, "color_large");

            buttonData_B.ToolTip = "Копировать связанное помещение у рассчитанного элемента отделки, либо связать с помещением напрямую";
            buttonData_B.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"https://www.youtube.com/watch?v=SR2YVR6bv8U&t=111s"));

            buttonData_C.ToolTip = "Подгрузка требуемых параметров в проект";
            buttonData_C.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"http://moodle/mod/book/view.php?id=502&chapterid=664#:~:text=%D0%B8%D0%BD%D1%81%D1%82%D1%80%D1%83%D0%BC%D0%B5%D0%BD%D1%82%D1%8B%20%D0%B0%D0%B2%D1%82%D0%BE%D0%BC%D0%B0%D1%82%D0%B8%D0%B7%D0%B0%D1%86%D0%B8%D0%B8%20R%D0%B5vit-,%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%9E%D0%A2%D0%94%D0%95%D0%9B%D0%9A%D0%90%22,-%D0%A1%D0%BA%D1%80%D0%B8%D0%BF%D1%82%20%D0%BF%D0%BE%20%D1%80%D0%B0%D1%81%D1%87%D0%B5%D1%82%D1%83"));

            buttonData_D.ToolTip = "Окрашивание элементов отделки по цветам согласно назначенному помещению";
            buttonData_D.LongDescription = "Необходимо пользоваться только на активном 3D виде";
            buttonData_D.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"https://www.youtube.com/watch?v=SR2YVR6bv8U&t=111s"));

            panel.AddStackedItems(buttonData_C, buttonData_B, buttonData_D);
            //UPDATE
            PushButtonData buttonData_E = new PushButtonData("Обновить", "Менеджер спецификаций", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.Update).FullName);
            PushButton button_E = panel.AddItem(buttonData_E) as PushButton;
            button_E.ToolTip = "Обновление спецификаций для общей ведомости проекта";
            button_E.LongDescription = string.Format("{0}", Names.assembly);
            button_E.ItemText = "Обновить\nведомость";
            button_E.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"https://www.youtube.com/watch?v=JB2h_qJFOjo"));
            SetIcon(button_E, "update_large");
            panel.AddSeparator();
            PushButtonData buttonData_F = new PushButtonData("Обновить группы", "Связать по типовому этажу", Assembly.GetExecutingAssembly().Location, typeof(ExternalCommands.ParseGroup).FullName);
            PushButton button_F = panel.AddItem(buttonData_F) as PushButton;
            button_F.ToolTip = "Инструмент позволяет привязать элементы отделки к помещениям в одной группе и попытаться найти аналогичные привязки в остальных экземплярах группы";
            button_F.LongDescription = string.Format("{0}", Names.assembly);
            button_F.ItemText = "Типовые\nэтажи";
            button_F.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, @"https://www.youtube.com/watch?v=2D8GMjI__pg"));
            SetIcon(button_F, "parse_groups");
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            return Result.Succeeded;
        }
        private void SetIcon(PushButton btn, string path)
        {
            BitmapImage icon = new BitmapImage(new Uri(string.Format(@"{0}\Source\{1}.png", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), path)));
            btn.Image = icon;
            btn.LargeImage = icon;
        }
        private void SetIcon(PushButtonData btn, string path)
        {
            BitmapImage icon = new BitmapImage(new Uri(string.Format(@"{0}\Source\{1}.png", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), path)));
            btn.Image = icon;
            btn.LargeImage = icon;
        }
        public void OnIdling(object sender, IdlingEventArgs args)
        {
            if (Preferences.IdlingCollectionsRun.Count != 0)
            {
                using (Transaction t = new Transaction((sender as UIApplication).ActiveUIDocument.Document, "KPLN Отделка"))
                {
                    t.Start();
                    while (Preferences.IdlingCollectionsRun.Count != 0)
                    {
                        AbstractOnIdlingCommand cmd = Preferences.IdlingCollectionsRun[0];
                        try
                        {
                            cmd.Execute(sender as UIApplication);
                            Preferences.IdlingCollectionsRun.RemoveAt(0);
                        }
                        catch (Exception e) { Output.PrintError(e); }

                    }
                    t.Commit();
                }
            }
        }
    }

}
