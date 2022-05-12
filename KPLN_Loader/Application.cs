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
        private const string _RibbonName = "KPLN";
        private string _RevitVersion;
        
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
            // Запуск основного процесса
            _RevitVersion = application.ControlledApplication.VersionNumber;
            Tools = new Tools_Environment(_RevitVersion);
            Tools.ClearPreviousLog();
            application.CreateRibbonTab(_RibbonName);
            string sqlPath = "Z:\\Отдел BIM\\03_Скрипты\\08_Базы данных\\KPLN_Loader.db";
            Print("Инициализация...", MessageType.Header);
            Logger.Trace($"---[Запуск в Revit {_RevitVersion}]---\n");
            if (File.Exists(sqlPath))
            {
                SQLiteDataBase = new Tools_SQL(sqlPath);
            }
            else
            {
                Print(
                    "Ошибка подключения БД. Будут загружены старые версии плагинов. Работа некоторых из них - будет ограничена!", 
                    MessageType.Error
                    );
                SQLiteDataBase = null;
            }

            // Запускается основной процесс инициализации
            if (SQLiteDataBase != null)
            {
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
            }
            
            // Обработка ошибки подключения БД
            else
            {
                DirectoryInfo userDirsRoot = Tools.ApplicationLocation;
                DirectoryInfo[] userDirs = userDirsRoot.GetDirectories();
                foreach (DirectoryInfo rootFolder in userDirs)
                {
                    if (rootFolder.Name.Contains(_RevitVersion))
                    {
                        userDirsRoot = rootFolder;
                    }
                }
                DirectoryInfo userRevitVersionDirs = new DirectoryInfo(Path.Combine(userDirsRoot.FullName, "Modules"));
                Logger.Info($"Ошибка подключения к БД. Осуществляю загрузку отсюда: {userRevitVersionDirs}\n");
                foreach(DirectoryInfo userFolder in userRevitVersionDirs.GetDirectories())
                {
                    try
                    {
                        // Загрузка модулей
                        UpdateModules(userFolder, application, null);
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }
                }
                Print("Готово!", MessageType.Header);
                Logger.Trace($"---[Успешная инициализация в Revit {_RevitVersion} без подключения к БД]---\n");

                return Result.Succeeded;
            }
            
            // Подключение к БД прошло успешно и в полном объеме
            if (User.Department != null)
            {
                try
                {
                    if (Tools.PrepareLocalDirectory())
                    {
                        List<SQLModuleInfo> foundedModules = new List<SQLModuleInfo>();
                        Print("Поиск доступных модулей...", MessageType.Header);
                        
                        //Загрузка общих модулей по отделу
                        foreach (SQLModuleInfo module in SQLiteDataBase.GetModules(User.Department.Id.ToString(), "Modules", _RevitVersion, "-1"))
                        {
                            foundedModules.Add(module);
                        }

                        //Загрузка общих модулей по отделу и проекту
                        foreach (SQLProjectInfo project in User_Projects)
                        {
                            foreach (SQLModuleInfo module in SQLiteDataBase.GetModules(User.Department.Id.ToString(), "Modules", _RevitVersion, project.Id.ToString()))
                            {
                                foundedModules.Add(module);
                            }
                        }
                        Logger.Info($"Успешное подключение к БД. Осуществляю загрузку отсюда: {sqlPath}\n");
                        foreach (SQLModuleInfo module in foundedModules)
                        {
                            Print(string.Format("Инфо: Загрузка модуля [{0}]", module.Name), MessageType.System_Regular);
                            try
                            {
                                DirectoryInfo loadedModule = Tools.CopyModuleFromPath(new DirectoryInfo(module.Path), module.Version, module.Name);
                                
                                // Загрузка модулей
                                UpdateModules(loadedModule, application, module);
                            }
                            catch (Exception e)
                            {
                                PrintError(e);
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
                Logger.Trace($"---[Успешная инициализация в Revit {_RevitVersion} с подключением к БД]---\n");
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

        private void UpdateModules(
            DirectoryInfo currentDir,
            UIControlledApplication application,
            SQLModuleInfo module
        )
        {
            foreach (FileInfo file in currentDir.GetFiles())
            {
                if (file.Name.Split('.').Last() == "dll")
                {
                    try
                    {
                        Logger.Info($"Загружаю dll отсюда {file.FullName}\n");
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

                        Result loadingResult = moduleInstance.Execute(application, _RibbonName);
                        if (loadingResult == Result.Succeeded && module != null)
                        {
                            LoadedModules.Add(moduleInstance);
                            Print(string.Format("Модуль [{0}] успешно активирован!", module.Name), MessageType.System_OK);
                        }
                        else if (module != null)
                        {
                            Print(string.Format("С модулем [{0}] есть проблемы!", module.Name), MessageType.Warning);
                        }
                    }
                    catch (Exception e)
                    {
                        PrintError(e);
                    }
                }
            }
        }

    }
}
