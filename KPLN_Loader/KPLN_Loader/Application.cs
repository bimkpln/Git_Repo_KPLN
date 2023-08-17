﻿using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Loader.Common;
using KPLN_Loader.Core.SQLiteData;
using KPLN_Loader.Forms;
using KPLN_Loader.Services;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Threading;

namespace KPLN_Loader
{
    public class Application : IExternalApplication
    {
        /// <summary>
        /// Queue для команд на выполнение по событию "OnIdling"
        /// </summary>
        public static Queue<IExecutableCommand> OnIdling_CommandQueue = new Queue<IExecutableCommand>();
        private readonly static string _diteTime = DateTime.Now.ToString("dd/MM/yyyy_HH/mm/ss");
        private readonly static string _sqlConfigPath = @"X:\BIM\5_Scripts\Git_Repo_KPLN\SQLConfig.json";
        private SQLiteService _dbService;
        private EnvironmentService _envService;
        private Logger _logger;
        private const string _ribbonName = "KPLN";
        private string _revitVersion;


        
        public delegate void LoadModulesProgress(int progress, string msg);
        public event LoadModulesProgress Progress;
        public UIControlledApplication App { get; private set; }



        public Result OnShutdown(UIControlledApplication application)
        {
            application.ControlledApplication.DocumentOpened -= new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);
            return Result.Succeeded;
        }

        /// <summary>
        /// Кэширование текщего пользователя из БД
        /// </summary>
        public static User CurrentRevitUser { get; private set; }

        public Result OnStartup(UIControlledApplication application)
        {
            ProgressForm progressForm = new ProgressForm(this, "asdas", 6);

            App = application;

            _revitVersion = application.ControlledApplication.VersionNumber;

            #region Настройка NLog
            LogManager.Setup().LoadConfigurationFromFile(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString() + "\\nlog.config");
            _logger = LogManager.GetLogger("KPLN_Loader");

            string logDirPath = $"c:\\temp\\KPLN_Logs\\{_revitVersion}";
            string logFileName = "KPLN_Loader";
            LogManager.Configuration.Variables["logdir"] = logDirPath;
            LogManager.Configuration.Variables["logfilename"] = logFileName;
            #endregion

            _logger.Info($"Запуск в Revit {_revitVersion}. Версия модуля: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");
            ClearingOldLogs(logDirPath, logFileName);
            try
            {
                #region Подготовка и проверка окружения
                _envService = new EnvironmentService(_logger, _revitVersion, _diteTime);
                _envService.SQLFilesExistChecker(_sqlConfigPath);
                _envService.PreparingAndCliningDirectories();
                #endregion

                #region Подготовка/создание пользователя
                string dbPath = _envService.DatabasesPaths.FirstOrDefault(d => d.Name.Contains("Main")).Path;
                _dbService = new SQLiteService(_logger, dbPath);
                CurrentRevitUser = _dbService.Authorization();
                #endregion

                // Создание панели Revit
                application.CreateRibbonTab(_ribbonName);

                #region Подготовка, копирование и активация модулей для пользователя
                _logger.Info("Подготовка, копирование и активация модулей для пользователя:");
                
                // Коллекция модулей из БД
                IEnumerable<Module> userAllModules = _dbService.GetModulesForCurrentUser(CurrentRevitUser);
                // Модули-библиотеки
                IEnumerable<Module> userLibModules = userAllModules.Where(m => m.IsLibraryModule.ToLower().Equals("true"));
                // Пользовательские загружаемые в ревит модули
                IEnumerable<Module> userModules = userAllModules.Where(m => m.IsLibraryModule.ToLower().Equals("false"));
                
                string userDescription = _dbService.GetDescriptionForCurrentUser(CurrentRevitUser);
                progressForm.Start();

                //try
                //{
                //    foreach (Module lib in userLibModules)
                //    {
                //        DirectoryInfo targetDirInfo = _envService.CopyModule(lib);
                //        if (targetDirInfo != null)
                //        {
                //            foreach (FileInfo file in targetDirInfo.GetFiles())
                //            {
                //                String[] spearator = { ".dll" };
                //                if (file.Name.Split(spearator, StringSplitOptions.None).Count() > 1)
                //                {
                //                    // Каждую dll библиотеки - нужно прогрузить в текущее приложение, чтобы появилась связь в проекте.
                //                    // Если этого не сделать - будут проблемы с использованием загружаемых библиотек
                //                    _ = System.Reflection.Assembly.LoadFrom(file.FullName);
                //                }
                //            }

                //            _logger.Info($"Модуль-библиотека [{lib.Name}] успешно скопирован!");
                //        }
                //        else
                //            _logger.Error($"Модуль-библиотека {lib.Name} - не скопировался!");

                //    }

                //    foreach (Module userModule in userModules)
                //    {
                //        DirectoryInfo targetDirInfo = _envService.CopyModule(userModule);
                //        if (targetDirInfo != null && ActivationModule(targetDirInfo, application, userModule))
                //        {
                //            //progressForm.Increment();
                //        }
                //        else
                //            _logger.Error($"Модуль {userModule.Name} - не скопировался!");
                //    }
                //}
                //catch (Exception ex)
                //{
                //    _logger.Error($"Локальная ошибка загрузки модуля: \n{ex}");
                //}
                #endregion
            }
            catch (Exception ex)
            {
                _logger.Error($"Глобальная ошибка плагина загрузки: \n{ex}");
                _logger.Info($"Инициализация не удалась\n");

                return Result.Cancelled;
            }

            _logger.Info($"Успешная инициализация в Revit {_revitVersion}\n");

            #region Подписка на события
            application.Idling += new EventHandler<IdlingEventArgs>(OnIdling);
            application.ControlledApplication.DocumentOpened += new EventHandler<DocumentOpenedEventArgs>(OnDocumentOpened);
            #endregion

            return Result.Succeeded;
        }

        public void DoWork()
        {
            int progress = 0;

            // Коллекция модулей из БД
            IEnumerable<Module> userAllModules = _dbService.GetModulesForCurrentUser(CurrentRevitUser);
            // Модули-библиотеки
            IEnumerable<Module> userLibModules = userAllModules.Where(m => m.IsLibraryModule.ToLower().Equals("true"));
            // Пользовательские загружаемые в ревит модули
            IEnumerable<Module> userModules = userAllModules.Where(m => m.IsLibraryModule.ToLower().Equals("false"));

            try
            {
                foreach (Module lib in userLibModules)
                {
                    DirectoryInfo targetDirInfo = _envService.CopyModule(lib);
                    if (targetDirInfo != null)
                    {
                        foreach (FileInfo file in targetDirInfo.GetFiles())
                        {
                            String[] spearator = { ".dll" };
                            if (file.Name.Split(spearator, StringSplitOptions.None).Count() > 1)
                            {
                                // Каждую dll библиотеки - нужно прогрузить в текущее приложение, чтобы появилась связь в проекте.
                                // Если этого не сделать - будут проблемы с использованием загружаемых библиотек
                                _ = System.Reflection.Assembly.LoadFrom(file.FullName);
                                progress++;
                                Progress?.Invoke(progress, $"Загружаю. Получил {progress} из {userAllModules.Count()} модулей");
                            }
                        }

                        _logger.Info($"Модуль-библиотека [{lib.Name}] успешно скопирован!");
                    }
                    else
                        _logger.Error($"Модуль-библиотека {lib.Name} - не скопировался!");

                }

                foreach (Module userModule in userModules)
                {
                    DirectoryInfo targetDirInfo = _envService.CopyModule(userModule);
                    if (targetDirInfo != null && ActivationModule(targetDirInfo, App, userModule))
                    {
                        progress++;
                        Progress?.Invoke(progress, $"Загружаю. Получил {progress} из {userAllModules.Count()} модулей");
                        //progressForm.Increment();
                    }
                    else
                        _logger.Error($"Модуль {userModule.Name} - не скопировался!");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"Локальная ошибка загрузки модуля: \n{ex}");
            }

        }

        /// <summary>
        /// Очистка от старых логов
        /// </summary>
        /// <param name="logPath">Путь к логам</param>
        private void ClearingOldLogs(string logPath, string logName)
        {
            if (Directory.Exists(logPath))
            {
                int cnt = 0;
                DirectoryInfo logDI = new DirectoryInfo(logPath);
                foreach (FileInfo log in logDI.EnumerateFiles())
                {
                    if (log.CreationTime.Date < DateTime.Now.AddDays(-5) && log.Name.Contains(logName))
                    {
                        try
                        {
                            log.Delete();
                            cnt++;
                        }
                        // Ошибка будет только если файл занят
                        catch (UnauthorizedAccessException) { }
                        catch (Exception ex)
                        {
                            _logger.Error($"При попытке очистки старых логов произошла ошибка: {ex.Message}");
                        }
                    }
                }
                _logger.Info($"Очистка от старых логов произведена успешно! Удалено {cnt}");
            }
        }

        /// <summary>
        /// Активация модуля в Revit
        /// </summary>
        /// <param name="currentDir">Путь к файлам</param>
        /// <param name="application">UIControlledApplication</param>
        /// <param name="userModule">Модуль для активации</param>
        private bool ActivationModule(DirectoryInfo currentDir, UIControlledApplication application, Module userModule)
        {
            foreach (FileInfo file in currentDir.GetFiles())
            {
                String[] spearator = { ".dll" };
                if (file.Name.Split(spearator, StringSplitOptions.None).Count() > 1)
                {
                    System.Reflection.Assembly moduleAssembly = System.Reflection.Assembly.LoadFrom(file.FullName);
                    Type implemnentationType = moduleAssembly.GetType(file.Name.Split(spearator, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() + ".Module", false);
                    if (implemnentationType != null)
                    {
                        implemnentationType.GetMember("Module");

                        IExternalModule moduleInstance = Activator.CreateInstance(implemnentationType) as IExternalModule;

                        Result loadingResult = moduleInstance.Execute(application, _ribbonName);
                        if (loadingResult == Result.Succeeded && userModule != null)
                        {
                            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(moduleAssembly.Location);
                            _logger.Info($"Модуль [{userModule.Name}] версии {fvi.FileVersion} успешно активирован!");
                            return true;
                        }
                        else if (userModule != null)
                        {
                            _logger.Error($"Модуль [{userModule.Name}] не найден!");
                        }
                        else if (loadingResult != Result.Succeeded)
                        {
                            _logger.Error($"Модуль [{userModule.Name}] не активирован. Проблема: \n{loadingResult}");
                        }
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Событие возникающее в режимах простоя Revit (когда приложению API становится безопасно обращаться к активному документу)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnIdling(object sender, IdlingEventArgs args)
        {
            UIApplication uiapp = sender as UIApplication;
            while (OnIdling_CommandQueue.Count != 0)
            {
                try
                {
                    OnIdling_CommandQueue.Dequeue().Execute(uiapp);
                }
                catch (Exception ex)
                {
                    TaskDialog td = new TaskDialog("ОШИБКА")
                    {
                        MainIcon = TaskDialogIcon.TaskDialogIconWarning,
                        MainInstruction = "Произошла серъезная ошибка при выполнении команды, обратись за помощью в BIM-отдел",
                        FooterText = $"См. файлы логов KPLN_Loader: c:\\temp\\KPLN_Logs\\{_revitVersion}"
                    };
                    td.Show();

                    _logger.Error($"Ошибка выполнения в событии OnIdling:\n{ex}");
                }
            }
        }

        /// <summary>
        /// Событие, происходящее при открытии документа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            Autodesk.Revit.ApplicationServices.Application app = sender as Autodesk.Revit.ApplicationServices.Application;
            _dbService.SetRevitUserName(app.Username, CurrentRevitUser);
        }
    }
}
