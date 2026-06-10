using KPLN_Library_ConfigWorker;
using KPLN_Tools.Forms.Models.Core;
using System.Diagnostics;
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
            HelpCommandCmd = new RelayCommand<object>(ExcecuteHelp);
        }

        public AutoSaveM CurrentAutoSaveM { get; set; }

        public ICommand OkCommandCmd { get; }

        public ICommand CloseWindowCmd { get; }

        public ICommand HelpCommandCmd { get; }

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

        private void ExcecuteHelp(object windObj)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = "chrome",
                Arguments = "http://moodle/mod/book/view.php?id=502&chapterid=1301#:~:text=%D0%9E%D0%A2%D0%94%D0%95%D0%9B%D0%AC%D0%9D%D0%AB%D0%99%20%D0%9F%D0%9B%D0%90%D0%93%D0%98%D0%9D%20%22%D0%90%D0%92%D0%A2%D0%9E%D0%A1%D0%9E%D0%A5%D0%A0%D0%90%D0%9D%D0%95%D0%9D%D0%98%D0%95%22", 
                UseShellExecute = true                
            };

            Process.Start(startInfo);


            CloseWindow(windObj);
        }
    }
}
