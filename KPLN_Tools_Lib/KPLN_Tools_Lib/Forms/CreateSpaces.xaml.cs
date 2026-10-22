using KPLN_Tools_Lib.Forms.Entities;
using KPLN_Tools_Lib.Forms.ViewModels;
using System.Windows;

namespace KPLN_Tools_Lib.Forms
{
    public partial class CreateSpaces : Window
    {
        internal CreateSpaces(CreateSpacesM model)
        {
            CurrentCreateSpacesVM = new CreateSpacesVM(this, model);
            InitializeComponent();
            DataContext = CurrentCreateSpacesVM;
        }

        internal CreateSpacesVM CurrentCreateSpacesVM { get; }
    }
}
