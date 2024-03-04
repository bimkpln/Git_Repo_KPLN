using KPLN_Library_Forms.Common;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    public partial class ElementPick : Window
    {
        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;

        private readonly IEnumerable<ElementEntity> _collection;

        public ElementPick(IEnumerable<ElementEntity> collection)
        {
            _collection = collection;
            InitializeComponent();

            Elements.ItemsSource = _collection;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        /// <summary>
        /// Выбранный элемент
        /// </summary>
        public ElementEntity SelectedElement { get; private set; }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public IEnumerable<ElementEntity> Collection { get { return _collection; } }

        /// <summary>
        /// Статус запуска
        /// </summary>
        public RunStatus Status { get; private set; }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Status = RunStatus.Close;
                Close();
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(SearchText);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_isRun)
            {
                Status = RunStatus.Close;
            }
        }

        private void OnElementClick(object sender, RoutedEventArgs e)
        {
            SelectedElement = (sender as Button).DataContext as ElementEntity;

            _isRun = true;
            Status = RunStatus.Run;
            Close();
        }

        /// <summary>
        /// Фильтрация по имени
        /// </summary>
        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string _searchName = textBox.Text.ToLower();

            ObservableCollection<ElementEntity> filteredElement = new ObservableCollection<ElementEntity>();

            foreach (ElementEntity elemnt in Collection)
            {
                if (elemnt.Name.ToLower().Contains(_searchName))
                {
                    filteredElement.Add(elemnt);
                }
            }

            Elements.ItemsSource = filteredElement;
        }

        /// <summary>
        /// Запуск без фильтрации
        /// </summary>
        private void OnRunWithoutFilterClick(object sender, RoutedEventArgs e)
        {
            SelectedElement = null;
            _isRun = true;
            Status = RunStatus.Run;
            Close();
        }
    }
}
