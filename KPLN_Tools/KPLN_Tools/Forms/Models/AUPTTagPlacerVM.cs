using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_Library_PluginActivityWorker;
using KPLN_Tools.ExternalCommands;
using KPLN_Tools.Forms.Models.Core;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms.Models
{
    public sealed class AUPTTagPlacerVM
    {
        private readonly string _cofigName = "AUPTTagPlacer";

        public AUPTTagPlacerVM(Document doc)
        {
            // Чтение конфигурации последнего запуска
            object configObj = ConfigService.ReadConfigFile<AUPTTagPlacerM>(ConfigType.Local, _cofigName);
            if (configObj != null && configObj is AUPTTagPlacerM model)
                AUPTTagPlacerModel = model;
            else
                AUPTTagPlacerModel = new AUPTTagPlacerM();

            AUPTTagPlacerModel.SetMainData(doc);

            RunCommandCmd = new RelayCommand<object>(ExecuteRun);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public AUPTTagPlacerM AUPTTagPlacerModel { get; set; }

        public ICommand RunCommandCmd { get; }

        public ICommand CloseWindowCmd { get; }

        private void ExecuteRun(object windObj)
        {            
            if (windObj is Window window)
            {
                if (string.IsNullOrEmpty(AUPTTagPlacerModel.SelectedTagTypeName))
                {
                    TaskDialog.Show(
                        ExtCmd_AUPT_TagPlacer.PluginName,
                        $"Перед запуском нужно выбрать типоразмер марки");

                    window.Activate();

                    return;
                }

                ConfigService.SaveConfig<AUPTTagPlacerM>(ConfigType.Local, AUPTTagPlacerModel, _cofigName);
                
                
                DBUpdater.UpdatePluginActivityAsync_ByPluginNameAndModuleName(ExtCmd_AUPT_TagPlacer.PluginName, ModuleData.ModuleName).ConfigureAwait(false);
                
                window.DialogResult = true;
                window.Close();
            }
        }

        private void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
