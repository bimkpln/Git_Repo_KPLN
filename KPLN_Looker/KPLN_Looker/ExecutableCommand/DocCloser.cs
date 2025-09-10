﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Loader.Common;
using System;
using System.Collections.Generic;
using System.Linq;
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

        private readonly DBUser _currentDBUser;
        
        private readonly Document _doc;
        private readonly string _docName;
        private readonly bool _sendToBitrix;

        public DocCloser(DBUser currentDBUser, Document doc, bool sendToBitrix = true)
        {
            _currentDBUser = currentDBUser;
            
            _doc = doc;
            _docName = _doc.Title;
            _sendToBitrix = sendToBitrix;
        }

        public Result Execute(UIApplication app)
        {
            app.DialogBoxShowing += App_DialogBoxShowing;

            UIDocument uidoc = app.ActiveUIDocument;

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
                    else
                    {
                        if (_sendToBitrix)
                            BitrixMessageSender
                                .SendMsg_ToBIMChat($"Проект {_docName} не удалось закрыть у пользователя {_currentDBUser.Surname} {_currentDBUser.Name}. " +
                                $"Проблема: \n{ex.Message}");

                        // Отписка в случае ошибки
                        app.DialogBoxShowing -= App_DialogBoxShowing;

                        return Result.Cancelled;
                    }
                }

            }
            catch (Exception ex)
            {
                if (_sendToBitrix)
                    BitrixMessageSender
                        .SendMsg_ToBIMChat($"Проект {_docName} не удалось закрыть у пользователя {_currentDBUser.Surname} {_currentDBUser.Name}. " +
                        $"Проблема: \n{ex.Message}");

                // Отписка в случае ошибки
                app.DialogBoxShowing -= App_DialogBoxShowing;

                return Result.Cancelled;
            }

            if (_sendToBitrix)
                BitrixMessageSender
                    .SendMsg_ToBIMChat($"Проект {_docName} у пользователя {_currentDBUser.Surname} {_currentDBUser.Name} успешно закрылся (если далее не следуют сообщения об ошибках).");

            return Result.Succeeded;
        }

        /// <summary>
        /// Событие открывающегося окна ревит
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void App_DialogBoxShowing(object sender, Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs args)
        {
            // Проверка открытого файла для самоотписки.
            UIApplication uiapp = sender as UIApplication;
            DocumentSet docSet = uiapp.Application.Documents;
            bool isDocClose = true;
            foreach (Document doc in docSet)
            {
                if (doc.Title == _docName)
                    isDocClose = false;
            }

            if (!isDocClose)
            {
                if (args.Cancellable)
                {
                    args.Cancel();
                }
                else
                {
                    DBRevitDialog currentDBDialog = DBMainService.DBRevitDialogColl.FirstOrDefault(rd => args.DialogId.Contains(rd.DialogId));
                    if (currentDBDialog != null)
                    {
                        if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                        {
                            bool overrideResult = args.OverrideResult((int)taskDialogResult);
                            if (!overrideResult && _sendToBitrix)
                                BitrixMessageSender.SendMsg_ToBIMChat($"Проект {_docName}: при закрывании проекта - не удалось переопределить окно {args.DialogId}. Необходим контроль со стороны человека");
                        }
                        else if (_sendToBitrix)
                            BitrixMessageSender.SendMsg_ToBIMChat($"Проект {_docName}: при закрывании проекта - не  удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!");
                    }
                    else if (_sendToBitrix)
                        BitrixMessageSender.SendMsg_ToBIMChat($"Проект {_docName}: при закрывании проекта - окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека");
                }
            }
            // Самоотписка после обработки окон, если проект закрылся (отписаться выше нельзя - обработчик событий работает ПОСЛЕ освобождения Idling, что НЕ приемлемо)
            else
                uiapp.DialogBoxShowing -= App_DialogBoxShowing;
        }

        static void CloseDocProc(object stateInfo) => SendKeys.SendWait("^{F4}");
    }
}
