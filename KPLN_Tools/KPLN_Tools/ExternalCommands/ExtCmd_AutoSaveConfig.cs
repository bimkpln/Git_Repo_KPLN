using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_DBWorker.FactoryParts.SQLite;
using KPLN_Tools.Forms;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KPLN_Tools.ExternalCommands
{
    /// <summary>
    /// Класс используется в модуле "KPLN_Looker" при обработке событий ViewActivated и DocumentSynchronizedWithCentral
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class ExtCmd_AutoSaveConfig : IExternalCommand
    {
        private const string _modelNameSepar = "_KAS_";
        
        internal const string PluginName = "Автосохранение";

        private static readonly Dictionary<string, DateTime> _lastAlarmByModel = new Dictionary<string, DateTime>();

        private static AutoSaveVM _aSVModelInst;
        private static SaveAsOptions _saveAsOptionsInst;

        internal static AutoSaveVM ASVModelInst
        {
            get
            {
                if (_aSVModelInst == null)
                    _aSVModelInst = new AutoSaveVM();

                return _aSVModelInst;
            }
            private set => _aSVModelInst = value;
        }

        internal static SaveAsOptions SaveAsOptionsInst
        {
            get
            {
                if (_saveAsOptionsInst == null)
                {
                    _saveAsOptionsInst = new SaveAsOptions()
                    {
                        OverwriteExistingFile = true,
                    };

                    WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions()
                    {
                        SaveAsCentral = false,
                    };

                    _saveAsOptionsInst.SetWorksharingOptions(worksharingSaveAsOptions);
                }

                return _saveAsOptionsInst;
            }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AutosaveForm autosaveForm = new AutosaveForm();
            if ((bool)autosaveForm.ShowDialog())
                ASVModelInst = autosaveForm.ASVModel;

            return Result.Succeeded;
        }

        /// <summary>
        /// Обновление счётчика последнего сохранения
        /// </summary>
        public static void UpdateLastAlarm(Document doc)
        {
            if (CheckDocError(doc))
                return;

            if (!ASVModelInst.CurrentAutoSaveM.IsAutoSaveEnabled)
                return;

            string modelKey_Title = Path.GetFileNameWithoutExtension(SQLiteDocService.GetFileFullName(doc));

            DateTime now = DateTime.Now;
            if (!_lastAlarmByModel.TryGetValue(modelKey_Title, out DateTime _))
                _lastAlarmByModel.Add(modelKey_Title, now);
            else
                _lastAlarmByModel[modelKey_Title] = now;
        }

        /// <summary>
        /// Сохранение локальной копии
        /// </summary>
        public static void Save(Document doc)
        {
            if (CheckDocError(doc))
                return;

            if (!ASVModelInst.CurrentAutoSaveM.IsAutoSaveEnabled)
                return;

            TimeSpan delayTime = new TimeSpan(0, ASVModelInst.CurrentAutoSaveM.SelectedInterval, 0);

            string modelKey_Title = Path.GetFileNameWithoutExtension(SQLiteDocService.GetFileFullName(doc));
            
            DateTime now = DateTime.Now;
            if (!_lastAlarmByModel.TryGetValue(modelKey_Title, out DateTime lastAlarm))
                _lastAlarmByModel.Add(modelKey_Title, now);
            else if (delayTime < (now - lastAlarm))
            {
                // Обнуляю счётчик (при возникновении ошибок всё равно обнулится,
                // это хорошо для стабильности, т.к. не уйдёт в бесконечный цикл,
                // это плохо для качества, т.к. при ошибке шаг сохранения будет пропущен) 
                _lastAlarmByModel[modelKey_Title] = now;


                // Взвожу флаг выполнения операции
                bool userSelect = true;
                if (ASVModelInst.CurrentAutoSaveM.ShowWarningWindow)
                {
                    var msgResult = MessageBox.Show(
                        ModuleData.MainWindowOwner,
                        "Выполнить сохранение?",
                        $"KPLN_{PluginName} [{modelKey_Title}]",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Asterisk);

                    userSelect = msgResult == DialogResult.Yes;
                }


                // Сохраняю, если юзер не передумал
                if (userSelect)
                {
                    if (ASVModelInst.CurrentAutoSaveM.IsSaveToCopiesEnabled)
                    {
                        string pathToSave = GetPathToLocalsSave();

                        if (string.IsNullOrEmpty(pathToSave))
                        {
                            MessageBox.Show(
                                ModuleData.MainWindowOwner,
                                "Ошибка, отправь разработчику! Не удалось определить путь к сохранению локальных копий.",
                                $"KPLN_{PluginName} [{modelKey_Title}]",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            return;
                        }

                        if (!Directory.Exists(pathToSave))
                        {
                            MessageBox.Show(
                                ModuleData.MainWindowOwner,
                                $"Ошибка, отправь разработчику! Путь для сохранения локальных копий - не существует: {pathToSave}",
                                $"KPLN_{PluginName} [{modelKey_Title}]",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            return;
                        }

                        string fullPath = Path.Combine(pathToSave, $"{modelKey_Title}{_modelNameSepar}{DateTime.Now:MMdd_HHmm}.rvt");
                        ModelPath mPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(fullPath);
                        try
                        {
                            doc.SaveAs(mPath, SaveAsOptionsInst);
                        }
                        catch (Exception ex)
                        {
                            // Ошибка открытой транзакции (например при работе с группой) - предупреждаю пользователя
                            if (ex.Message.Contains("Unable to close all open transaction phases!"))
                            {
                                // Переписываю значение, чтобы не удваивать шаг при такого типа ошибке,
                                // а чтобы он сработал раньше - через 2 мин
                                _lastAlarmByModel[modelKey_Title] = now - new TimeSpan(0, ASVModelInst.CurrentAutoSaveM.SelectedInterval, -120);

                                return;
                            }


                            MessageBox.Show(
                                ModuleData.MainWindowOwner,
                                $"Ошибка, отправь разработчику! Не удалось сохранить, ошибка:\n\n{ex.Message}. \n\nПуть к модели для сохранения: {fullPath}",
                                $"KPLN_{PluginName} [{modelKey_Title}]",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                            return;
                        }


                        DeleteBackupFiles(modelKey_Title, ASVModelInst.CurrentAutoSaveM.CopiesCountSelected, pathToSave);
                    }
                    else
                        doc.Save();
                }
            }
        }

        private static bool CheckDocError(Document doc)
        {
            if (doc == null
                || !doc.IsWorkshared
                || doc.IsDetached
                || doc.IsFamilyDocument)
                return true;

            return false;
        }

        private static string GetPathToLocalsSave()
        {
            string iniFile = GetPathToINIFile();

            // Читаем секцию [Directories], ключ UserDataDir
            string userDataDir = null;
            bool inSection = false;

            foreach (var line in File.ReadAllLines(iniFile))
            {
                if (line.Trim() == "[Directories]") 
                { 
                    inSection = true; 
                    continue; 
                }
                if (inSection && line.StartsWith("[")) 
                    break; // следующая секция
                if (inSection && line.StartsWith("ProjectPath="))
                {
                    userDataDir = line.Substring("ProjectPath=".Length).Trim();
                    break;
                }
            }

            return userDataDir != null ? Environment.ExpandEnvironmentVariables(userDataDir) : null;
        }

        private static string GetPathToINIFile()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Autodesk",
                "Revit",
                $"Autodesk Revit {ModuleData.RevitVersion}",
                "Revit.ini"
            );
        }

        private static void DeleteBackupFiles(string modelKey_Title, int archFilesLimit, string pathTo)
        {
            string clearTaskMsg = string.Empty;
            Task clearArchCopyFilesTask = Task.Run(() =>
            {
                string[] archFiles = Directory.GetFiles(pathTo);
                FileInfo[] currentCentralArchCopies = archFiles
                    .Where(a => a.Contains(modelKey_Title) && a.Contains(_modelNameSepar))
                    .Select(a => new FileInfo(a))
                    .OrderBy(fi => fi.CreationTime)
                    .ToArray();

                if (currentCentralArchCopies.Count() <= archFilesLimit)
                    return;

                int startCount = currentCentralArchCopies.Count() - archFilesLimit;
                while (startCount > 0)
                {
                    startCount--;
                    FileInfo archCopyToDel = currentCentralArchCopies[startCount];
                    archCopyToDel.Delete();
                }
            });
            Task.WaitAll(new Task[1] { clearArchCopyFilesTask });

            if (clearTaskMsg != string.Empty)
                MessageBox.Show(
                    ModuleData.MainWindowOwner,
                    $"Отправь разработчику! Ошибка при очистке старых резервных копий: {clearTaskMsg}",
                    $"KPLN_{PluginName} [{modelKey_Title}]",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
        }
    }
}
