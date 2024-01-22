using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_BIMTools_Ribbon.Common;
using NLog;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.ExternalCommands
{
    /// <summary>
    /// Плагин по импорту моделей с Revit-Server
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandImportFromRS : ReviServerEnvironment, IExternalCommand, ICopier
    {
        public CommandImportFromRS() : base()
        {
        }
        
        public CommandImportFromRS(UIControlledApplication application, Logger logger) : base(application, logger)
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

            Logger.Info($"Старт импорта файлов");
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
                if (OpenAndCopyFile(app, openOptions, saveAsOptions, filePath, FilePathesTo.FirstOrDefault()))
                    CountProcessedDocs++;
            }

            SendResultMsg("Плагин по импорту моделей с Revit-Server", FilePathesFrom.Count());
            Logger.Info($"Файлы успешно импортированы!");

            //Отписка от событий
            RevitUIControlledApp.DialogBoxShowing -= OnDialogBoxShowing;
            RevitUIControlledApp.ControlledApplication.DocumentOpened -= OnDocumentOpened;
            RevitUIControlledApp.ControlledApplication.DocumentClosed -= OnDocumentClosed;
            //RevitUIControlledApp.ControlledApplication.FailuresProcessing -= OnFailureProcessing;

            return Result.Succeeded;
        }

        public void SetPathes(string pathToConfig)
        {
            FilePathesFrom = new List<string>()
            {
                "RSN://192.168.0.5/Речной порт Якутск/Стадия_Р/АР/РПЯ_К1_РД_АР_R23.rvt",
            };

            FilePathesTo = new List<string>()
            {
                "Y:\\Жилые здания\\Речной порт Якутск\\10.Стадия_Р\\5.АР\\1.RVT\\RevitServer_Архив"
            };
        }
    }
}
