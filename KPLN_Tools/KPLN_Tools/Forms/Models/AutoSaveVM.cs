using KPLN_Library_ConfigWorker;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools.Forms.Models
{
    public sealed class AutoSaveVM
    {
        private readonly string _cofigName = "AutoSaveConfig";

        public AutoSaveVM()
        {
            // Чтение конфигурации последнего запуска
            object configObj = ConfigService.ReadConfigFile<AutoSaveM>(ConfigType.Local, _cofigName);
            if (configObj != null && configObj is AutoSaveM model)
                CurrentAutoSaveM = model;
            else
                CurrentAutoSaveM = new AutoSaveM();

            OkCommandCmd = new RelayCommand<object>(ExecuteOk);
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public AutoSaveM CurrentAutoSaveM { get; set; }

        public ICommand OkCommandCmd { get; }

        public ICommand CloseWindowCmd { get; }

        private void ExecuteOk(object windObj)
        {
            ConfigService.SaveConfig<AutoSaveM>(ConfigType.Local, CurrentAutoSaveM, _cofigName);
            if (windObj is Window window)
            {
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
