using Autodesk.Revit.DB;
using KPLN_Publication.ExternalCommands.PublicationSet;
using System.Collections.Generic;
using System.Windows;

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
                KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new CommandCreateSet(Elements, tb.Text));
                Close();
            }

        }
    }
}
