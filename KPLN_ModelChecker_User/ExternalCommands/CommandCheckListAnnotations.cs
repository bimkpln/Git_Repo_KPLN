using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_ModelChecker_Lib.Commands;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms;
using System.Linq;
using System.Windows.Input;
using System.Windows.Interop;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public sealed class CommandCheckListAnnotations : AbstrCommand, IExternalCommand
    {
        public CommandCheckListAnnotations() : base() { }
        public static string ConfigName = "CheckListAnnotationsSettingsConfig";
        public static ConfigType ConfigType = ConfigType.Local;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;


            // Чтение конфигурации последнего запуска
            CheckListAnnotationsSettingsM settingsM;
            object lastRunConfigObj = ConfigService.ReadConfigFile<CheckListAnnotationsSettingsM>(ConfigType, ConfigName);
            if (lastRunConfigObj != null && lastRunConfigObj is CheckListAnnotationsSettingsM settings)
                settingsM = settings;
            else
                settingsM = new CheckListAnnotationsSettingsM();


            // Проверка на зажатый Ctrl для вызова формы настроек
            bool configurationRequested = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
            if (configurationRequested)
            {
                CheckListAnnotationsSettingsForm form = new CheckListAnnotationsSettingsForm(settingsM);

                new WindowInteropHelper(form)
                {
                    Owner = uiapp.MainWindowHandle
                };

                form.ShowDialog();


                return Result.Cancelled;
            }


            // Основной функционал команды
            CommandCheck = new CheckListAnnotations().Set_UIAppData(uiapp, uiapp.ActiveUIDocument.Document);
            ElemsToCheck = CommandCheck.GetElemsToCheck();

            if (ElemsToCheck.Count() > 0)
                ExecuteByUIApp<CheckListAnnotations>(uiapp, settingsM);

            return Result.Succeeded;
        }
    }
}
