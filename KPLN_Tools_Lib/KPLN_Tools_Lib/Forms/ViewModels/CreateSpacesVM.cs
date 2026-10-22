using KPLN_Tools_Lib.Forms.Commands;
using KPLN_Tools_Lib.Forms.Entities;
using System.Windows;
using System.Windows.Input;

namespace KPLN_Tools_Lib.Forms.ViewModels
{
    internal sealed class CreateSpacesVM
    {
        internal CreateSpacesVM(Window mainWindow, CreateSpacesM model)
        {
            CurrentCreateSpacesM = model;
            RunCmd = new RelayCommand(() => mainWindow.DialogResult = true, () => model.CanRun);
            CloseWindowCmd = new RelayCommand(() => mainWindow.Close());
            SelectAllCmd = new RelayCommand(() => SetSelection(true));
            SelectNoneCmd = new RelayCommand(() => SetSelection(false));
        }

        public CreateSpacesM CurrentCreateSpacesM { get; }
        public ICommand RunCmd { get; }
        public ICommand CloseWindowCmd { get; }
        public ICommand SelectAllCmd { get; }
        public ICommand SelectNoneCmd { get; }

        private void SetSelection(bool isSelected)
        {
            foreach (CreateSpacesLinkM link in CurrentCreateSpacesM.Links)
                link.IsSelected = isSelected;
        }
    }
}
