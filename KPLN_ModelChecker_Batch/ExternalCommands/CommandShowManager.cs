using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ModelChecker_Batch.Forms;
using NLog;

namespace KPLN_ModelChecker_Batch.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CommandShowManager : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            #region Настройка NLog
            // Конфиг для логгера лежит в KPLN_Loader. Это связано с инициализацией dll самим ревитом. Настройку тоже производить в основном конфиге
            Logger currentLogger = LogManager.GetLogger("KPLN_BIMTools");

            string logDirPath = $"c:\\KPLN_Temp\\KPLN_Logs\\{ModuleData.RevitVersion}";
            string logFileName = "KPLN_BIMTools";
            LogManager.Configuration.Variables["bimtools_logdir"] = logDirPath;
            LogManager.Configuration.Variables["bimtools_logfilename"] = logFileName;
            #endregion

            UIApplication uiapp = commandData.Application;

            BatchManager mainForm = new BatchManager(currentLogger, uiapp);
            if (!(bool)mainForm.ShowDialog())
                return Result.Cancelled;

            return Result.Succeeded;
        }
    }
}
