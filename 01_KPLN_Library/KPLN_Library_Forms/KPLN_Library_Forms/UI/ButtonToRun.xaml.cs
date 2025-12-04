using KPLN_Library_Forms.Common;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class ButtonToRun : Window
    {
        private readonly string _mainDescription = string.Empty;

        private readonly IEnumerable<ButtonToRunEntity> _collection;

        public ButtonToRun(string mainTitle, IEnumerable<ButtonToRunEntity> collection)
        {
            _mainDescription = mainTitle;
            _collection = collection;
            InitializeComponent();

            BtnColl.ItemsSource = _collection;
            MainDescription.Text = _mainDescription;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public IEnumerable<ButtonToRunEntity> Collection { get { return _collection; } }

        /// <summary>
        /// Нажатая кнопка
        /// </summary>
        public ButtonToRunEntity SelectedButton { get; private set; }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void OnBtnClick(object sender, RoutedEventArgs e)
        {
            SelectedButton = (sender as Button).DataContext as ButtonToRunEntity;

            DialogResult = true;
            Close();
        }
    }
}
