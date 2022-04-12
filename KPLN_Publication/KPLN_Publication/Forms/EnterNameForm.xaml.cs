using Autodesk.Revit.DB;
using KPLN_Publication.ExternalCommands.BeforePublication;
using KPLN_Publication.ExternalCommands.Print;
using KPLN_Publication.ExternalCommands.PublicationSet;
using KPLN_Publication.Common;
using System;
using System.Collections.Generic;
using System.Linq;
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

namespace KPLN_Publication.Forms
{
    /// <summary>
    /// Логика взаимодействия для EnterNameForm.xaml
    /// </summary>
    public partial class EnterNameForm : Window
    {
        private List<View> Elements { get; set; }
        public EnterNameForm(List<View> elements, Window parent)
        {
            Owner = parent;
            Elements = elements;
            InitializeComponent();
        }

        private void OnApply(object sender, RoutedEventArgs e)
        {
            if (tb.Text == string.Empty || string.IsNullOrWhiteSpace(tb.Text))
            {
                Close();
            }
            else
            {
                KPLN_Loader.Preferences.CommandQueue.Enqueue(new CommandCreateSet(Elements, tb.Text));
                Close();
            }
            
        }
    }
}
