using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using static KPLN_NavisWorksReports.Common.Collections;

namespace KPLN_NavisWorksReports.Forms
{
    public class KPTaskDialog
    {
        public KPTaskDialogResult DialogResult = KPTaskDialogResult.None;
        private ConfirmDialog DialogWindow { get; set; }
        public bool IsShown = false;
        public KPTaskDialog(Window parent, string title, string header, string message, KPTaskDialogIcon iconType, bool canCancel, string footer = null)
        {
            if (parent == null) { throw new Exception("Parent window can not be null!"); }
            if (header == null) { throw new Exception("Header can not be null!"); }
            if (message == null) { throw new Exception("Message can not be null!"); }
            DialogWindow = new ConfirmDialog(parent, this, title, header, message, iconType, canCancel, footer);
        }
        public KPTaskDialogResult ShowDialog()
        {
            if (!IsShown)
            {
                DialogWindow.ShowDialog();
                IsShown = true;
            }
            else
            {
                throw new Exception("The dialog has already been shown!");
            }
            return DialogResult;
        }
    }

}
