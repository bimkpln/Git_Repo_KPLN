extern alias revit;
using KPLN_ModelChecker_Coordinator.DB;
using KPLN_ModelChecker_Coordinator.Tools;
using KPLN_Library_DataBase.Collections;
using KPLN_Loader.Common;
using revit.Autodesk.Revit.DB;
using revit.Autodesk.Revit.DB.Events;
using revit.Autodesk.Revit.UI;
using revit.Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using static KPLN_ModelChecker_Coordinator.ModuleData;
using static KPLN_Loader.Output.Output;
using System.Windows.Interop;
using System.Windows;
using KPLN_ModelChecker_Coordinator.Common;

namespace KPLN_ModelChecker_Coordinator
{
    public class Module : IExternalModule
    {
        private readonly string _assemblyPath = Assembly.GetExecutingAssembly().Location;

        public Result Close()
        {
            return Result.Succeeded;
        }
        public Result Execute(UIControlledApplication application, string tabName)
        {
#if Revit2020
            MainWindowHandle = application.MainWindowHandle;
            HwndSource hwndSource = HwndSource.FromHwnd(MainWindowHandle);
            RevitWindow = hwndSource.RootVisual as Window;
#endif
#if Revit2018
            try
            {
                MainWindowHandle = WindowHandleSearch.MainWindowHandle.Handle;
            }
            catch (Exception) { }
#endif

            //Добавляю панель
            RibbonPanel currentPanel = application.CreateRibbonPanel(tabName, "Проверки");
            
            //Добавляю кнопки в панель
            if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
            {
                AddPushButtonDataInPanel(
                    "Открыть менеджер проверок", 
                    "Серийная\nпроверка", 
                    "Запуск проверки выбранных документов на ошибки.",
                    string.Format(
                        "Выполняет анализ выбранных моделей на предмет наличия ошибок по требованиям регламента 'BIM-контроль документации' и подсчитывает их." +
                            "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExternalCommands.CommandOpenDialog).FullName,
                    currentPanel,
                    "icon_manager.png",
                    "http://moodle.stinproject.local"
                );
            }

            AddPushButtonDataInPanel(
                    "Открыть окно статистики",
                    "Окно\ncтатистики",
                    "Отображение статистики по документам, которые были проверены на ошибки.",
                    string.Format(
                        "Выводит количественные показатели ошибок в модели по требованиям регламента 'BIM-контроль документации'." +
                            "\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExternalCommands.CommandShowStatistics).FullName,
                    currentPanel,
                    "icon_browser.png",
                    "http://moodle.stinproject.local"
                );

            currentPanel.AddSlideOut();
            
            if (KPLN_Loader.Preferences.User.Department.Id == 4 || KPLN_Loader.Preferences.User.Department.Id == 6)
            {
                AddPushButtonDataInPanel(
                    "Параметры",
                    "Параметры",
                    "Редактирование пользовательских настроек.",
                    string.Format("\nДата сборки: {0}\nНомер сборки: {1}\nИмя модуля: {2}",
                        ModuleData.Date,
                        ModuleData.Version,
                        ModuleData.ModuleName
                    ),
                    typeof(ExternalCommands.CommandShowSettings).FullName,
                    currentPanel,
                    "icon_setup.png",
                    "http://moodle.stinproject.local"
                );
            }
            application.DialogBoxShowing += OnDialogBoxShowing;
            application.ControlledApplication.ApplicationInitialized += OnInitialized;
            application.ControlledApplication.FailuresProcessing += OnFailureProcessing;
            application.ControlledApplication.DocumentOpened += OnOpened;
            return Result.Succeeded;
        }

        private void CheckDocument(Document doc, DbDocument dbDoc)
        {
            try
            {
                DbRowData rowData = new DbRowData();
                rowData.Errors.Add(new DbError("Ошибка привязки к уровню", CheckTools.CheckLevels(doc)));
                rowData.Errors.Add(new DbError("Зеркальные элементы", CheckTools.CheckMirrored(doc)));
                rowData.Errors.Add(new DbError("Ошибка мониторинга осей", CheckTools.CheckMonitorGrids(doc)));
                rowData.Errors.Add(new DbError("Ошибка мониторинга уровней", CheckTools.CheckMonitorLevels(doc)));
                rowData.Errors.Add(new DbError("Ошибки семейств и экземпляров", CheckTools.CheckFamilies(doc)));
                rowData.Errors.Add(new DbError("Ошибки подгруженных связей", CheckTools.CheckSharedLocations(doc) + CheckTools.CheckLinkWorkSets(doc)));
                rowData.Errors.Add(new DbError("Предупреждения Revit", CheckTools.CheckErrors(doc)));
                rowData.Errors.Add(new DbError("Размер файла", CheckTools.CheckFileSize(dbDoc.Path)));
                rowData.Errors.Add(new DbError("Элементы в наборах подгруженных связей", CheckTools.CheckElementsWorksets(doc)));
                DbController.WriteValue(dbDoc.Id.ToString(), rowData.ToString());
                //BotActions.SendRegularMessage(string.Format("👌 @{0}_{1} завершил проверку документа #{3} #{2}", NormalizeString(KPLN_Loader.Preferences.User.Family), NormalizeString(KPLN_Loader.Preferences.User.Name), NormalizeString(dbDoc.Project.Name), NormalizeString(dbDoc.Name)), Bot.Target.Process);
            }
            catch (Exception) { }
        }
       
        public void OnOpened(object sender, DocumentOpenedEventArgs args)
        {
            try
            {
                Document doc = args.Document;
                if (!CheckTools.AllWorksetsAreOpened(doc)) { return; }
                if (doc.IsWorkshared && !doc.IsDetached)
                {
                    string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath());
                    FileInfo centralPath = new FileInfo(path);
                    foreach (DbDocument dbDoc in KPLN_Library_DataBase.DbControll.Documents)
                    {
                        FileInfo metaCentralPath = new FileInfo(dbDoc.Path);
                        if (dbDoc.Code == "NONE") { continue; }
                        if (centralPath.FullName == metaCentralPath.FullName)
                        {
                            if (File.Exists(string.Format(@"Z{0}\doc_id_{1}.sqlite", MyPathes.ModelCheckerDBPath, dbDoc.Id.ToString())))
                            {
                                List<DbRowData> rows = DbController.GetRows(dbDoc.Id.ToString());
                                if (rows.Count != 0)
                                {
                                    if ((DateTime.Now.Day - rows.Last().DateTime.Day > 14 && rows.Last().DateTime.Day != DateTime.Now.Day) || rows.Last().DateTime.Month != DateTime.Now.Month || rows.Last().DateTime.Year != DateTime.Now.Year)
                                    {
                                        CheckDocument(doc, dbDoc);
                                    }
                                }
                                else
                                {
                                    CheckDocument(doc, dbDoc);
                                }
                            }
                            else
                            {
                                CheckDocument(doc, dbDoc);
                            }
                        }
                    }
                }
            }
            catch (Exception) { }
        }
        
        public void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            if (!ModuleData.AutoConfirmEnabled || !ModuleData.up_close_dialogs)
            {
                return;
            }
            bool gotErrors = false;
            foreach (FailureMessageAccessor i in args.GetFailuresAccessor().GetFailureMessages())
            {
                if (i.GetSeverity() == FailureSeverity.Warning)
                {
                    args.GetFailuresAccessor().DeleteWarning(i);
                }
                else
                {
                    args.GetFailuresAccessor().ResolveFailure(i);
                    gotErrors = true;
                }
                args.GetFailuresAccessor().DeleteAllWarnings();
                args.GetFailuresAccessor().ResolveFailures(args.GetFailuresAccessor().GetFailureMessages());
            }
            if (gotErrors)
            {
                args.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
            }
            else
            {
                args.SetProcessingResult(FailureProcessingResult.Continue);
            }
        }
        
        public void OnInitialized(object sender, ApplicationInitializedEventArgs args)
        {
            KPLN_Library_DataBase.DbControll.Update();
        }

        public void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            if (!ModuleData.AutoConfirmEnabled || !ModuleData.up_close_dialogs)
            {
                return;
            }
            
            HashSet<string> dialogIds = new HashSet<string>();
            SQLiteConnection sql = new SQLiteConnection();
            try
            {
                sql.ConnectionString = MyPathes.MainDBConnection;
                sql.Open();
                using (SQLiteCommand cmd = new SQLiteCommand("SELECT Name FROM TaskDialogs", sql))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            dialogIds.Add(rdr.GetString(0));
                        }
                    }
                }
                sql.Close();
            }
            catch (Exception)
            {
                try
                {
                    sql.Close();
                }
                catch (Exception) { }
            }
            
            if (dialogIds.Contains(args.DialogId))
            {
                Print(string.Format("Всплывающий диалог: [{0}]...", args.DialogId), KPLN_Loader.Preferences.MessageType.System_Regular);
                sql = new SQLiteConnection();
                string value = "NONE";
                
                try
                {
                    sql.ConnectionString = MyPathes.MainDBConnection;
                    sql.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(string.Format("SELECT OverrideResult FROM TaskDialogs WHERE Name = '{0}'", args.DialogId), sql))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                value = rdr.GetString(0);
                            }
                        }
                    }
                    sql.Close();
                }
                catch (Exception)
                {
                    try
                    {
                        sql.Close();
                    }
                    catch (Exception) { }
                }
                
                if (value != "NONE")
                {
                    TaskDialogResult rslt;
                    if (Enum.TryParse(value, out rslt))
                    {
                        args.OverrideResult((int)rslt);
                        Print(string.Format("...действие «по умолчанию» - [{0}]", rslt.ToString("G")), KPLN_Loader.Preferences.MessageType.System_Regular);
                    }
                    else
                    {
                        Print($"Окно не обработалось {value}", KPLN_Loader.Preferences.MessageType.Regular);
                    }

                }
                else
                {
                    Print("...действие «по умолчанию» - [NONE]", KPLN_Loader.Preferences.MessageType.System_Regular);
                }
            }
            else
            {
                Print(string.Format("Не удалось идентифицировать всплывающий диалог: [{0}]", args.DialogId), KPLN_Loader.Preferences.MessageType.Critical);
                
                try
                {
                    sql.ConnectionString = MyPathes.MainDBConnection;
                    sql.Open();
                    using (SQLiteCommand cmd = sql.CreateCommand())
                    {
                        cmd.CommandText = "INSERT INTO TaskDialogs ([Name], [OverrideResult]) VALUES (@Name, @OverrideResult)";
                        cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@Name", Value = args.DialogId });
                        cmd.Parameters.Add(new SQLiteParameter() { ParameterName = "@OverrideResult", Value = "NONE" });
                        cmd.ExecuteNonQuery();
                    }
                    sql.Close();
                }
                catch (Exception)
                {
                    try
                    {
                        sql.Close();
                    }
                    catch (Exception) { }
                }
                
                try
                {
                    Thread t = new Thread(() =>
                    {
                        SaveImage(args.DialogId.ToString());
                    });
                    t.Start();
                }
                catch (Exception) { }
                
                try
                {
                    if (args.Cancellable)
                    {
                        args.Cancel();
                        Print("...действие «по умолчанию» - [Закрыть]", KPLN_Loader.Preferences.MessageType.System_Regular);
                    }
                }
                catch (Exception) { }
            }
        }

        /// <summary>
        /// Метод для добавления отдельной в панель
        /// </summary>
        /// <param name="name">Внутреннее имя кнопки</param>
        /// <param name="text">Имя, видимое пользователю</param>
        /// <param name="shortDescription">Краткое описание, видимое пользователю</param>
        /// <param name="longDescription">Полное описание, видимое пользователю при залержке курсора</param>
        /// <param name="className">Имя класса, содержащего реализацию команды</param>
        /// <param name="panel">Панель, в которую добавляем кнопку</param>
        /// <param name="imageName">Имя иконки</param>
        /// <param name="contextualHelp">Ссылка на web-страницу по клавише F1</param>
        private void AddPushButtonDataInPanel(string name, string text, string shortDescription, string longDescription, string className, RibbonPanel panel, string imageName, string contextualHelp)
        {
            PushButtonData data = new PushButtonData(name, text, _assemblyPath, className);
            PushButton button = panel.AddItem(data) as PushButton;
            button.ToolTip = shortDescription;
            button.LongDescription = longDescription;
            button.ItemText = text;
            BtnImagine(button, imageName);
            button.SetContextualHelp(new ContextualHelp(ContextualHelpType.Url, contextualHelp));
            button.AvailabilityClassName = typeof(Availability.StaticAvailable).FullName;
        }

        /// <summary>
        /// Метод для добавления иконки для кнопки
        /// </summary>
        /// <param name="button">Кнопка, куда нужно добавить иконку</param>
        /// <param name="imageName">Имя иконки с раширением</param>
        private void BtnImagine(RibbonButton button, string imageName)
        {
            string imageFullPath = Path.Combine(new FileInfo(_assemblyPath).DirectoryName, @"Imagens\", imageName);
            button.LargeImage = new BitmapImage(new Uri(imageFullPath));

        }

        /// <summary>
        /// Сохранение скрина монитора с ошибкой
        /// </summary>
        /// <param name="name"></param>
        private static void SaveImage(string name)
        {
            try
            {
                Thread.Sleep(1500);
                int screenLeft = SystemInformation.VirtualScreen.Left;
                int screenTop = SystemInformation.VirtualScreen.Top;
                int screenWidth = SystemInformation.VirtualScreen.Width;
                int screenHeight = SystemInformation.VirtualScreen.Height;
                Bitmap bitmap = new Bitmap(screenWidth, screenHeight);
                Graphics graphics = Graphics.FromImage(bitmap as Image);
                graphics.CopyFromScreen(screenLeft, screenTop, 0, 0, bitmap.Size);
                bitmap.Save(string.Format(@"{0}\{1}_{2}.jpg", MyPathes.ModelCheckerCommonPath, name, Guid.NewGuid().ToString()), ImageFormat.Jpeg);
            }
            catch (Exception)
            { }
        }
    }
}
