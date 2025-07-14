using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_OpeningHoleManager.Core;
using KPLN_OpeningHoleManager.ExecutableCommand;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;

namespace KPLN_OpeningHoleManager.Forms
{
    public partial class ParamsOpenHoleByIOSElemsForm : Window
    {
        private static readonly Regex _regex = new Regex(@"[^0-9.,]+");

        public ParamsOpenHoleByIOSElemsForm(MVVMCore_MainMenu.MainViewModel currentVM)
        {
            ParamsOpenHoleByIOSElemsForm_VM = currentVM;
            
            InitializeComponent();
            DataContext = ParamsOpenHoleByIOSElemsForm_VM;

            PreviewKeyDown += new KeyEventHandler(HandleKeyboardPush);
        }

        /// <summary>
        /// Ссылка на ViewModel
        /// </summary>
        public MVVMCore_MainMenu.MainViewModel ParamsOpenHoleByIOSElemsForm_VM { get; private set; }

        private void HandleKeyboardPush(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();

            if (e.Key == Key.Enter)
            {
                this.DialogResult = true;
                SetConfigData();
                this.Close();
            }
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

        private void OnBtnApply_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            SetConfigData();
            this.Close();
        }

        private void SetConfigData() => MainConfig.SetData_ToConfig(ParamsOpenHoleByIOSElemsForm_VM);
    }
}
