using Autodesk.Revit.UI;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class MainMenu : System.Windows.Controls.Page, IDockablePaneProvider
    {
        private static readonly Regex _regex = new Regex(@"[^0-9.,]+");

        public MainMenu()
        {
            InitializeComponent();
            DataContext = new MVVMCore_MainMenu.ViewModel();
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Left
            };
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = _regex.IsMatch(e.Text);
        }

        private void TextBox_Pasting(object sender, DataObjectPastingEventArgs e)
        {
            if (e.DataObject.GetDataPresent(typeof(string)))
            {
                string text = (string)e.DataObject.GetData(typeof(string));
                if (_regex.IsMatch(text))
                    e.CancelCommand();
            }
            else
                e.CancelCommand();
        }
    }
}
