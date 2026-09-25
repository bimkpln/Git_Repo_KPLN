using KPLN_Tools_OVVK.Forms.Models;
using System.Windows;

namespace KPLN_Tools_OVVK.Forms
{
    public partial class OVVK_ScheduleSubForm : Window
    {
        public OVVK_ScheduleSubForm(string header, string currentValue)
        {
            InitializeComponent();

            DataContext = new ScheduleSubFormVM(header, currentValue);
        }

        public ScheduleSubFormVM SSFVm => DataContext as ScheduleSubFormVM;
    }
}
