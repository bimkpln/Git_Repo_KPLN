using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_CoordiantorAI.Common;
using KPLN_CoordiantorAI.Forms;
using System;
using System.Windows;
using System.Windows.Interop;

namespace KPLN_CoordiantorAI.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class StartAI : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                CoordinatorAiRepository repository = new CoordinatorAiRepository();

                if (!repository.DatabaseExists)
                {
                    message = "БД KPLN_CoordinatorAI не найдена. Обратитесь в BIM-отдел.";
                    TaskDialog.Show("Координатор ИИ", message);
                    return Result.Failed;
                }

                repository.EnsureDatabase();

                CurrentUserContext userContext = new CurrentUserContextService().GetCurrentUserContext();
                ChatWindow chatWindow = new ChatWindow(repository, new GigaChatClient(), userContext);
                SetRevitOwner(chatWindow);
                chatWindow.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ошибка", ex.ToString());
                return Result.Failed;
            }
        }

        private static void SetRevitOwner(Window window)
        {
            if (window == null || ModuleData.RevitMainWindowHandle == IntPtr.Zero)
                return;

            WindowInteropHelper helper = new WindowInteropHelper(window);
            helper.Owner = ModuleData.RevitMainWindowHandle;
        }
    }
}