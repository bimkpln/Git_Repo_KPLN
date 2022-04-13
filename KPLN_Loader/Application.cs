using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using static KPLN_Loader.Preferences;
using static KPLN_Loader.Output.Output;
using KPLN_Loader.Common;
using System;
using System.Linq;
using System.IO;
using System.Reflection;
using KPLN_Loader.Forms;
using Autodesk.Revit.DB.Events;
using System.Collections.Generic;
using System.Windows.Interop;
using System.Windows;

namespace KPLN_Loader
{
    public class Application : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public Result OnStartup(UIControlledApplication application)
        {
#if Revit2020 || Revit2022
            MainWindowHandle = application.MainWindowHandle;
            HwndSource hwndSource = HwndSource.FromHwnd(MainWindowHandle);
            RevitWindow = hwndSource.RootVisual as Window;
#endif
#if Revit2018
            try
            {
                MainWindowHandle = SystemTools.WindowHandleSearch.MainWindowHandle.Handle;
            }
            catch (Exception e)
            {
                PrintError(e);
            }
#endif
            Print("Инициализация...", MessageType.Header);
            Tools = new Tools_Environment(application.ControlledApplication.VersionNumber);
            Tools.ClearPreviousLog();
            assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString();
            application.CreateRibbonTab("KPLN");
            SQLiteDataBase = new Tools_SQL();
            if (SQLiteDataBase.IfUserExist(System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last()))
            {
                SQLiteDataBase.GetUserData(System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last());
                Print(string.Format("Добро пожаловать, {0}!", User.Name), MessageType.Success);
                try { SQLiteDataBase.GetUserProjects(User.SystemName, true); }
                catch (Exception e) { PrintError(e); }
                
            }
            else
            {
                WPFLogIn logInForm = new WPFLogIn();
                bool wasHiden = false;
                try
                {
                    if (FormOutput.Visible == true)
                    {
                        FormOutput.Hide();
                        wasHiden = true;
                    }
                }
                catch (Exception) { }
                logInForm.cbxDepartment.ItemsSource = SQLiteDataBase.GetDepartments();
                logInForm.ShowDialog();
                if (wasHiden) { FormOutput.Show(); }
                SQLiteDataBase.GetUserData(System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split('\\').Last());
                Print(string.Format("Добро пожаловать, {0}!", User.Name), MessageType.Success);
            }
            if (User.Department != null)
            {
                try
                {
                    if (Tools.PrepareLocalDirectory())
                    {
                        List<SQLModuleInfo> foundedModules = new List<SQLModuleInfo>();
                        Print("Поиск доступных модулей...", MessageType.Header);
                        foreach (SQLModuleInfo module in SQLiteDataBase.GetModules(User.Department.Id.ToString(), "Modules", application.ControlledApplication.VersionNumber, "-1"))
                        {
                            foundedModules.Add(module);
                        }
                        foreach (SQLProjectInfo project in User_Projects)
                        {
                            foreach (SQLModuleInfo module in SQLiteDataBase.GetModules(User.Department.Id.ToString(), "Modules", application.ControlledApplication.VersionNumber, project.Id.ToString()))
                            {
                                foundedModules.Add(module);
                            }
                        }
                        foreach (SQLModuleInfo module in foundedModules)
                        {
                            Print(string.Format("Инфо: Загрузка модуля [{0}]", module.Name), MessageType.System_Regular);
                            try
                            {
                                DirectoryInfo loadedModule = Tools.CopyModuleFromPath(new DirectoryInfo(module.Path), module.Version, module.Name);
                                foreach (FileInfo file in loadedModule.GetFiles())
                                {
                                    if (file.Name.Split('.').Last() == "dll")
                                    {
                                        try
                                        {
                                            Assembly assembly = Assembly.LoadFrom(file.FullName);
                                            Type implemnentationType = assembly.GetType(file.Name.Split('.').First() + ".Module");
                                            try
                                            {
                                                implemnentationType.GetMember("Module");
                                            }
                                            catch (Exception)
                                            {
                                                continue;
                                            }
                                            IExternalModule moduleInstance = Activator.CreateInstance(implemnentationType) as IExternalModule;
                                            
                                            Result loadingResult = moduleInstance.Execute(application, "KPLN");
                                            if (loadingResult == Result.Succeeded)
                                            {
                                                loadedModules.Add(moduleInstance);
                                                Print(string.Format("Модуль [{0}] успешно активирован!", module.Name), MessageType.System_OK);
                                            }
                                            else
                                            {
                                                Print(string.Format("С модулем [{0}] есть проблемы!", module.Name), MessageType.Warning);
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            PrintError(e, string.Format("Ошибка при получении модуля «{0}»", module.Name));
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                PrintError(e, string.Format("Ошибка при получении модуля «{0}»", module.Name));
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    PrintError(e);
                }
                //register events
                application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
                application.ApplicationClosing += new EventHandler<ApplicationClosingEventArgs>(OnClosing);
                application.ControlledApplication.DocumentSynchronizingWithCentral += new EventHandler<DocumentSynchronizingWithCentralEventArgs>(OnDocumentSynchronizing);
                //end
                Print("Готово!", MessageType.Header);
                UpdateUserInfo();
                return Result.Succeeded;
            }
            return Result.Failed;
        }
        public void OnDocumentSynchronizing(object sender, DocumentSynchronizingWithCentralEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIControlledApplication controlledApplication = sender as UIControlledApplication;
            UpdateUserInfo();
        }
        public void OnIdling(object sender, IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            UIControlledApplication controlledApplication = sender as UIControlledApplication;
            while (CommandQueue.Count != 0)
            {
                try
                {
                    CommandQueue.Dequeue().Execute(uiapp);
                }
                catch (Exception e) { PrintError(e); }    
            }
        }
        public void OnClosing(object sender, ApplicationClosingEventArgs args)
        {
            try { FormOutput.Close(); }
            catch (Exception) { }
        }
        private void UpdateUserInfo()
        {
            if (User != null)
            {
                try
                {
                    SQLiteDataBase.UpdateUserConnection(User.SystemName, "Users");
                    SQLiteDataBase.GetUserData(User.SystemName);
                    SQLiteDataBase.GetUserProjects(User.SystemName, false);
                    SQLiteDataBase.GetUsers();
                }
                catch (Exception)
                { }
            }
        }
    }
}
