using KPLN_Tools_Lib.Forms.Entities;
using KPLN_Tools_Lib.Forms.ViewModels;
using System.Windows;

namespace KPLN_Tools_Lib.Forms
{
    public partial class TagSpaces : Window
    {
        internal TagSpaces(TagSpacesM model)
        {
            CurrentTagSpacesVM = new TagSpacesVM(this, model);
            InitializeComponent();
            DataContext = CurrentTagSpacesVM;
        }

        internal TagSpacesVM CurrentTagSpacesVM { get; }
    }
}
