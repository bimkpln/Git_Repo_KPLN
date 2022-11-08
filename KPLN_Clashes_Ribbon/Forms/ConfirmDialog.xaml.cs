using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static KPLN_Clashes_Ribbon.Common.Collections;

namespace KPLN_Clashes_Ribbon.Forms
{
    /// <summary>
    /// Логика взаимодействия для ConfirmDialog.xaml
    /// </summary>
    public partial class ConfirmDialog : Window
    {
        private KPTaskDialog Dialog { get; }

        public ConfirmDialog(Window parent, KPTaskDialog taskDialog, string title, string header, string message, KPTaskDialogIcon iconType, bool canCandel, string footer = null)
        {
            Title = "KPLN: " + title;
            Dialog = taskDialog;
            try
            {
                Owner = parent;
            }
            catch (Exception) { }
            InitializeComponent();
            tbHeader.Text = header;
            tbBody.Text = message;
            if (!canCandel) { btnCancel.Visibility = Visibility.Collapsed; }
            switch (iconType)
            {
                case KPTaskDialogIcon.Happy: tbIcon.Text = ":)"; break;
                case KPTaskDialogIcon.Sad: tbIcon.Text = ":("; break;
                case KPTaskDialogIcon.Question: tbIcon.Text = "??"; break;
                case KPTaskDialogIcon.Warning: tbIcon.Text = "!!"; break;
                case KPTaskDialogIcon.Lol: tbIcon.Text = ":D"; break;
                case KPTaskDialogIcon.Ooo: tbIcon.Text = ":O"; break;
                default: tbIcon.Text = "#"; break;
            }
            if (footer != null && footer != string.Empty)
            {
                tbFooter.Text = footer;
            }
            else
            {
                tbFooter.Visibility = Visibility.Collapsed;
            }
            if (iconType == KPTaskDialogIcon.Warning || iconType == KPTaskDialogIcon.Ooo)
            {
                SystemSounds.Beep.Play();
            }
            else
            {
                SystemSounds.Hand.Play();
            }
        }
        private void OnCancel(object sender, RoutedEventArgs e)
        {
            Dialog.DialogResult = Common.Collections.KPTaskDialogResult.Cancel;
            Close();
        }

        private void OnOk(object sender, RoutedEventArgs e)
        {
            Dialog.DialogResult = Common.Collections.KPTaskDialogResult.Ok;
            Close();
        }
    }
}
