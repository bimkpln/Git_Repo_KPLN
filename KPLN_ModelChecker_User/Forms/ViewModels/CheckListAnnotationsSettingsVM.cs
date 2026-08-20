using Autodesk.Revit.UI;
using KPLN_Library_ConfigWorker;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.ExternalCommands;
using KPLN_ModelChecker_User.Forms.Commands;
using System.Windows;
using System.Windows.Input;
using MediaColor = System.Windows.Media.Color;
using WinForms = System.Windows.Forms;

namespace KPLN_ModelChecker_User.Forms.ViewModels
{
    public sealed class CheckListAnnotationsSettingsVM
    {
        private readonly CheckListAnnotationsSettingsForm _mainWindow;

        public CheckListAnnotationsSettingsVM(CheckListAnnotationsSettingsForm mainWindow, CheckListAnnotationsSettingsM settingsM)
        {
            _mainWindow = mainWindow;
            CurrentCheckListAnnotationsSettingsM = settingsM;

            // Реализация команд для кнопок окна
            ChooseColorCmd = new RelayCommand<object>(_ => ChooseColor());
            ClearHighlightCmd = new RelayCommand<object>(_ => ClearHighlight());
            SaveWindowCmd = new RelayCommand<object>(_ => SaveWindow());
        }

        public CheckListAnnotationsSettingsM CurrentCheckListAnnotationsSettingsM { get; set; }

        /// <summary>
        /// Команда для кнопки выбора цвета
        /// </summary>
        public ICommand ChooseColorCmd { get; }

        /// <summary>
        /// Команда для кнопки очистки подсветки
        /// </summary>
        public ICommand ClearHighlightCmd { get; }

        /// <summary>
        /// Команда для кнопки сохранения настроек и закрытия окна
        /// </summary>
        public ICommand SaveWindowCmd { get; }

        /// <summary>
        /// Команда для кнопки закрытия окна без сохранения
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void ChooseColor()
        {
            MediaColor color = CurrentCheckListAnnotationsSettingsM.HighlightColor;

            using (WinForms.ColorDialog dialog = new WinForms.ColorDialog())
            {
                dialog.FullOpen = true;
                dialog.Color = System.Drawing.Color.FromArgb(color.R, color.G, color.B);

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    CurrentCheckListAnnotationsSettingsM.HighlightColor = MediaColor.FromRgb(
                        dialog.Color.R,
                        dialog.Color.G,
                        dialog.Color.B);
                }
            }
        }

        public void ClearHighlight()
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new ExcCmdClearCurrentHighlight());

            _mainWindow.DialogResult = true;
            _mainWindow.Close();
        }

        public void SaveWindow()
        {
            ConfigService.SaveConfig<CheckListAnnotationsSettingsM>(CommandCheckListAnnotations.ConfigType, CurrentCheckListAnnotationsSettingsM, CommandCheckListAnnotations.ConfigName);

            _mainWindow.DialogResult = true;
            _mainWindow.Close();
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }
    }
}
