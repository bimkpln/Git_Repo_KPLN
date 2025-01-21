using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using System;
using KPLN_HoleManager.Forms;

namespace KPLN_HoleManager.Common
{
    // Регитсрация DockablePane
    internal static class DockablePreferences
    {
        public static DockableManagerForm Page = new DockableManagerForm();
        public static Guid PageGuid = new Guid("42246bf5-7ea2-4ce9-94ef-61e87d352a4c");

        public static void RegisterDockablePane(UIControlledApplication application)
        {
            DockableManagerForm page = new DockableManagerForm();

            DockablePaneProviderData data = new DockablePaneProviderData
            {
                FrameworkElement = page,
                InitialState = new DockablePaneState
                {
                }
            };

            application.RegisterDockablePane(new DockablePaneId(PageGuid), "Менеджер отверстий", page);

            // Скрываем панель после полной загрузки приложения
            application.ControlledApplication.ApplicationInitialized += (sender, args) =>
            {
                try
                {
                    DockablePane pane = application.GetDockablePane(new DockablePaneId(PageGuid));
                    if (pane.IsShown())
                    {
                        pane.Hide();
                    }
                }
                catch{} // Тут ничего не происходит
            };
        }
    }

    // Вызов DockablePane при нажатии на кнопку
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowDockablePane : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
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
                }

                return Result.Succeeded;
            }

            System.Windows.Forms.MessageBox.Show("Не удалось открыть панель", "Ошибка",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);

            return Result.Failed;
        }
    }
}