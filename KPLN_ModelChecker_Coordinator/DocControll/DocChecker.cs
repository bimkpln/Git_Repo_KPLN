extern alias revit;
using revit.Autodesk.Revit.DB;
using revit.Autodesk.Revit.UI;
using KPLN_ModelChecker_Coordinator.DB;
using KPLN_ModelChecker_Coordinator.Tools;
using KPLN_Loader.Common;
using KPLN_Library_DataBase.Collections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using static KPLN_Loader.Output.Output;
using System.Windows.Forms;
using System.Runtime.Remoting.Channels;
using System.Threading;
using revit::Autodesk.Revit.DB.Events;
using revit::Autodesk.Revit.UI.Events;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;
using System.Windows.Interop;
using System.Windows;
using KPLN_ModelChecker_Coordinator.SystemTools;

namespace KPLN_ModelChecker_Coordinator.DocControll
{
    public static class UiInput
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        public static void KeyEnter(Object stateInfo)
        {
            try
            {
                IntPtr current = GetForegroundWindow();
                foreach (Process process in Process.GetProcessesByName("Revit"))
                {
                    try
                    {
                        SetForegroundWindow(process.Handle);
                        SendKeys.SendWait("{ENTER}");
                        KeyEnterLoop(process.Handle);
                    }
                    catch (Exception)
                    { }
                }
                SetForegroundWindow(current);
            }
            catch (Exception)
            { }
            //SetForegroundWindow(Process.GetCurrentProcess().Handle);
            //SendKeys.SendWait("{ENTER}");
        }
        private static void KeyEnterLoop(IntPtr handle)
        {
            try
            {
                foreach (IntPtr childHandle in new WindowHandleInfo(handle).GetAllChildHandles())
                {
                    try
                    {
                        SetForegroundWindow(childHandle);
                        SendKeys.SendWait("{ENTER}");
                        KeyEnterLoop(childHandle);
                    }
                    catch (Exception)
                    { }
                }
            }
            catch (Exception)
            { }
        }
    }
    public class DocChecker : IExecutableCommand
    {
        public DocChecker(List<DbDocument> documents)
        {
            Documents = documents;
        }
        private List<DbDocument> Documents { get; }
        private static System.Threading.Timer _Timer { get; set; }
        private static string NormalizeString(string value)
        {
            string result = string.Empty;
            foreach (char c in value)
            {
                if (char.IsWhiteSpace(c)) { result += "_"; continue; }
                if (char.IsPunctuation(c)) { result += "_"; continue; }
                if (char.IsLetter(c) || char.IsDigit(c)) { result += c; continue; }
                continue;
            }
            return result;
        }
        public Result Execute(UIApplication app)
        {
            if (ModuleData.up_send_enter)
            {
                Thread thread = new Thread(() =>
                {
                    var autoEvent = new AutoResetEvent(true);
                    _Timer = new System.Threading.Timer(UiInput.KeyEnter, autoEvent, 20000, 20000);
                });
                thread.Start();
            }
            try
            {
                Print(string.Format("[{0}] Запуск...", DateTime.Now.ToString("T")), KPLN_Loader.Preferences.MessageType.Header);
                Print("Создание [placeholder] документа...", KPLN_Loader.Preferences.MessageType.Regular);
                
                string assemblyPath = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;
                
                ModuleData.AutoConfirmEnabled = true;
                
                app.Application.NewProjectDocument(UnitSystem.Metric);
                
                foreach (DbDocument doc in Documents)
                {
                    try
                    {
                        Print(string.Format("[{0}] Открытие {1}...", DateTime.Now.ToString("T"), doc.Path), KPLN_Loader.Preferences.MessageType.Header);
                        ModelPath path = ModelPathUtils.ConvertUserVisiblePathToModelPath(doc.Path);
                        OpenOptions options = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets, Audit = false };
                        WorksetConfiguration config = new WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets);
                        options.SetOpenWorksetsConfiguration(config);
                        Document docu = app.Application.OpenDocumentFile(path, options);
                        
                        try
                        {
                            DbRowData rowData = new DbRowData();
                            rowData.Errors.Add(new DbError("Ошибка привязки к уровню", CheckTools.CheckLevels(docu)));
                            rowData.Errors.Add(new DbError("Зеркальные элементы", CheckTools.CheckMirrored(docu)));
                            rowData.Errors.Add(new DbError("Ошибка мониторинга осей", CheckTools.CheckMonitorGrids(docu)));
                            rowData.Errors.Add(new DbError("Ошибка мониторинга уровней", CheckTools.CheckMonitorLevels(docu)));
                            rowData.Errors.Add(new DbError("Ошибки семейств и экземпляров", CheckTools.CheckFamilies(docu)));
                            rowData.Errors.Add(new DbError("Ошибки подгруженных связей", CheckTools.CheckSharedLocations(docu) + CheckTools.CheckLinkWorkSets(docu)));
                            rowData.Errors.Add(new DbError("Предупреждения Revit", CheckTools.CheckErrors(docu)));
                            rowData.Errors.Add(new DbError("Размер файла", CheckTools.CheckFileSize(doc.Path)));
                            rowData.Errors.Add(new DbError("Элементы в наборах подгруженных связей", CheckTools.CheckElementsWorksets(docu)));
                            
                            DbController.WriteValue(doc.Id.ToString(), rowData.ToString());
                        }
                        catch (Exception e)
                        {
                            DebugTools.Error(e);
                        }
                        
                        Print(string.Format("[{0}] Закрытие документа...", DateTime.Now.ToString("T")), KPLN_Loader.Preferences.MessageType.Header);
                        docu.Close(false);
                    }
                    catch (Exception e)
                    {
                        DebugTools.Error(e);
                    }
                }
                ModuleData.AutoConfirmEnabled = false;
                
                if (_Timer != null)
                {
                    _Timer.Dispose();
                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                ModuleData.AutoConfirmEnabled = false;
                
                DebugTools.Error(e);
                
                if (_Timer != null)
                {
                    _Timer.Dispose();
                }

                return Result.Failed;
            }
        }

    }
}
