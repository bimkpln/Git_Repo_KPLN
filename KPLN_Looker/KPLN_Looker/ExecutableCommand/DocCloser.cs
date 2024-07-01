using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace KPLN_Looker.ExecutableCommand
{
    internal class DocCloser : IExecutableCommand
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private bool _isMsgShowed = false;
        private string _docName;

        public DocCloser()
        {
        }

        public Result Execute(UIApplication app)
        {
            app.DialogBoxShowing += App_DialogBoxShowing;
            
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;
            _docName = doc.Title;
            
            IList<UIView> openedViewsColl = uidoc.GetOpenUIViews();
            try
            {
                foreach (UIView view in openedViewsColl)
                {
                    view.Close();
                }
            }
            // Закрыть все окна открытого проекта - не позволяет Revit API. В этом случае посылаю Ctrl+F4, т.е. закрываю последнюю вкладку с клавы
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
            {
                if (Module.MainWindowHandle != IntPtr.Zero)
                {
                    // Установить окно в качестве активного
                    SetForegroundWindow(Module.MainWindowHandle);
                    
                    // Закрыть его через Ctrl+F4
                    if (ex.Message.IndexOf("Cannot close a project's only open view.", StringComparison.OrdinalIgnoreCase) >= 0)
                        ThreadPool.QueueUserWorkItem(new WaitCallback(CloseDocProc));
                }

                BitrixMessageSender
                    .SendMsg_ToBIMChat($"После синхронизации проекта {_docName} - он успешно закрылся.");
            }
            catch (Exception ex)
            {
                if (!_isMsgShowed)
                    BitrixMessageSender
                        .SendMsg_ToBIMChat($"После синхронизации проекта {_docName} - его не удалось закрыть. Проблема: \n{ex.Message}");
            }
            finally
            {
                app.DialogBoxShowing -= App_DialogBoxShowing;
            }

            return Result.Succeeded;
        }

        private void App_DialogBoxShowing(object sender, Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs e)
        {
            _isMsgShowed = true;
            BitrixMessageSender
                .SendMsg_ToBIMChat($"После синхронизации проекта {_docName} - открылось окно {e.DialogId}. Скорее всего сам проект не закрылся.");
        }

        static void CloseDocProc(object stateInfo) => SendKeys.SendWait("^{F4}");
    }
}
