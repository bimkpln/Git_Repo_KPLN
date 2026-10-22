using KPLN_Tools_Lib.Forms.Commands;
using KPLN_Tools_Lib.Forms.Entities;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools_Lib.Forms.ViewModels
{
    internal sealed class TagSpacesVM
    {
        internal TagSpacesVM(Window mainWindow, TagSpacesM model)
        {
            CurrentTagSpacesM = model;
            RunCmd = new RelayCommand(() => mainWindow.DialogResult = true, () => model.CanRun);
            CloseWindowCmd = new RelayCommand(() => mainWindow.Close());
        }

        public TagSpacesM CurrentTagSpacesM { get; }
        public ICommand RunCmd { get; }
        public ICommand CloseWindowCmd { get; }
    }
}
