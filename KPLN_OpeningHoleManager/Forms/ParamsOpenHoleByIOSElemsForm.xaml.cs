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
            
            currentVM.RefreshTypeMapping();
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

            if (e.Key == Key.Enter && !(e.OriginalSource is System.Windows.Controls.ComboBox))
            {
                ApplySettings();
                e.Handled = true;
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
            ApplySettings();
        }

        private void OnAutoDetect_Click(object sender, RoutedEventArgs e)
        {
            ParamsOpenHoleByIOSElemsForm_VM.RefreshTypeMapping();
            ParamsOpenHoleByIOSElemsForm_VM.TypeMapping.AutoDetect();
        }

        private void OnAutoDetectEmpty_Click(object sender, RoutedEventArgs e)
        {
            // Сначала очищаем назначения, которых больше нет в текущей модели.
            ParamsOpenHoleByIOSElemsForm_VM.RefreshTypeMapping();
            ParamsOpenHoleByIOSElemsForm_VM.TypeMapping.AutoDetect(true);
        }

        private void OnLoadFamilies_Click(object sender, RoutedEventArgs e)
        {
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(
                new LoadOpeningFamiliesExcCmd(ParamsOpenHoleByIOSElemsForm_VM, Module.CurrentUIApplication.ActiveUIDocument.Document));
            Close();
        }

        private void ApplySettings()
        {
            try
            {
                MainConfig.SetData_ToConfig(ParamsOpenHoleByIOSElemsForm_VM);
                DialogResult = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(this, "Не удалось сохранить настройки.\n" + KPLN_OpeningHoleManager.Services.OpeningUserMessage.FromException(ex),
                    "Настройки отверстий", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
