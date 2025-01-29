using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System;
using KPLN_HoleManager.Forms;
using KPLN_HoleManager.ExternalCommand;

namespace KPLN_HoleManager.Common
{
    // Регитсрация DockablePane
    public static class DockablePreferences
    {
        public static DockableManagerForm Page;
        public static Guid PageGuid = new Guid("42246bf5-7ea2-4ce9-94ef-61e87d352a4c");

        public static void EnsureDockablePaneRegistered(UIControlledApplication application)
        {
            if (Page != null) {return;}

            Page = new DockableManagerForm();

            DockablePaneProviderData data = new DockablePaneProviderData
            {
                FrameworkElement = Page,
                InitialState = new DockablePaneState()
                {
                    DockPosition = DockPosition.Floating
                }
            };

            application.RegisterDockablePane(new DockablePaneId(PageGuid), "Менеджер отверстий", Page);
        }

        // Метод закрытия панелим при открытии нового документа
        public static void HideDockablePane(UIApplication uiApplication)
        {
            DockablePaneId paneId = new DockablePaneId(PageGuid);

            try
            {
                DockablePane pane = uiApplication.GetDockablePane(paneId);
                if (pane != null && pane.IsShown())
                {
                    pane.Hide();
                }
            }
            catch {}
        }
    }

    // Вызов DockablePane при нажатии на кнопку
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowDockablePane : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                _ExternalEventHandler.Initialize();

                // Получаем DockablePane
                DockablePaneId paneId = new DockablePaneId(DockablePreferences.PageGuid);
                DockablePane pane = commandData.Application.GetDockablePane(paneId);

                if (pane != null)
                {
                    if (pane.IsShown())
                    {
                        pane.Hide();
                    }
                    else
                    {
                        pane.Show();
                        DockablePreferences.Page.SetUIApplication(commandData.Application); // Передаём текущую сессию UI
                    }

                    return Result.Succeeded;
                }

                message = "Панель не найдена.";
            }
            catch (Exception ex)
            {
                message = $"Ошибка при работе с панелью: {ex.Message}";
            }

            return Result.Failed;
        }
    }
}