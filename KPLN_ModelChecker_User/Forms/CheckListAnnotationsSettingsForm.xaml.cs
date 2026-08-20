using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.Forms.ViewModels;
using System.Windows;

namespace KPLN_ModelChecker_User.Forms
{
    public partial class CheckListAnnotationsSettingsForm : Window
    {
        public CheckListAnnotationsSettingsForm(CheckListAnnotationsSettingsM settingsM)
        {
            CurrentCheckListAnnotationsSettingsVM = new CheckListAnnotationsSettingsVM(this, settingsM);

            InitializeComponent();

            DataContext = CurrentCheckListAnnotationsSettingsVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public CheckListAnnotationsSettingsVM CurrentCheckListAnnotationsSettingsVM { get; set; }
    }
}
