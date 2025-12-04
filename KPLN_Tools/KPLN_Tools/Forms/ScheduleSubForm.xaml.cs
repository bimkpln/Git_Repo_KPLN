using KPLN_Tools.Forms.Models;
using System.Windows;

namespace KPLN_Tools.Forms
{
    public partial class ScheduleSubForm : Window
    {
        public ScheduleSubForm(string header, string currentValue)
        {
            InitializeComponent();

            DataContext = new ScheduleSubFormVM(header, currentValue); ;
        }

        public ScheduleSubFormVM SSFVm => DataContext as ScheduleSubFormVM;
    }
}
