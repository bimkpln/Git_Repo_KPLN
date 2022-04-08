using KPLN_DataBase.Collections;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace KPLN_NavisWorksReports.Forms
{
    public class ProjectPickDialog
    {
        public static DbProject PickedProject { get; set; }
        private ProjectPicker form { get; set; }
        public ProjectPickDialog(Window parent)
        {
            PickedProject = null;
            form = new ProjectPicker(parent);
        }
        public void ShowDialog()
        {
            form.ShowDialog();
        }
        public bool IsConfirmed()
        {

            if (PickedProject == null) { return false; }
            return true;
        }
        public DbProject GetLastPickedProject()
        {
            return PickedProject;
        }
    }
}
