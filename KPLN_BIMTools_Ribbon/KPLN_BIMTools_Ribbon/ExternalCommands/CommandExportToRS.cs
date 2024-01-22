using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using KPLN_Library_SQLiteWorker.FactoryParts;
using NLog;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    /// <summary>
    /// Плагин по экспорту моделей на Revit-Server
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandExportToRS : ReviServerEnvironment, IExternalCommand, ICopier
    {
        public CommandExportToRS() : base()
        {
        }
        
        public CommandExportToRS(UIControlledApplication application, Logger logger) : base(application, logger)
        {
        }

        public List<string> FilePathesFrom { get; set; }

        public List<string> FilePathesTo { get; set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Подписка на события
            RevitUIControlledApp.DialogBoxShowing += OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened += OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed += OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing += OnFailureProcessing;

            Logger.Info($"Старт экспорта файлов");
            SetPathes("blas");

            // Подготовка к открытию
            Application app = commandData.Application.Application;
            OpenOptions openOptions = new OpenOptions() { DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets };
            openOptions.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));

            SaveAsOptions saveAsOptions = new SaveAsOptions() { OverwriteExistingFile = true };
            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions() { SaveAsCentral = true };
            saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);

            // Копирую файлы по указанным путям
            foreach (string filePath in FilePathesFrom)
            {
                OpenAndCopyFile(app, openOptions, saveAsOptions, filePath, FilePathesTo.FirstOrDefault());
            }

            SendResultMsg("Плагин по экспорту моделей на Revit-Server", 0);
            Logger.Info($"Файлы успешно экспортированы!");
            
            //Отписка от событий
            RevitUIControlledApp.DialogBoxShowing -= OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed -= OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= OnFailureProcessing;


            return Result.Succeeded;
        }

        public void SetPathes(string pathToConfig)
        {
            throw new System.NotImplementedException();
        }
    }
}
