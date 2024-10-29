using KPLN_Library_Forms.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using static KPLN_Library_Forms.Common.UIStatus;

namespace KPLN_Library_Forms.UI
{
    public partial class ElementSinglePick : Window
    {
        /// <summary>
        /// Флаг для идентификации запуска приложения, а не закрытия через Х (любое закрытие окна связано с Window_Closing, поэтому нужен доп. флаг)
        /// </summary>
        private bool _isRun = false;

        private readonly IEnumerable<ElementEntity> _collection;
        private readonly ObservableCollection<ElementEntity> _showCollection;

        public ElementSinglePick(IEnumerable<ElementEntity> collection)
        {
            _collection = collection;
            _showCollection = new ObservableCollection<ElementEntity>(_collection);
            InitializeComponent();

            this.Title = $"KPLN: Выбери нужный элемент";
            Elements.ItemsSource = _showCollection;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        public ElementSinglePick(IEnumerable<ElementEntity> collection, string title) : this(collection)
        {
            this.Title = $"KPLN: {title}";
        }

        /// <summary>
        /// Выбранный элемент
        /// </summary>
        public ElementEntity SelectedElement { get; private set; }

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

            foreach (ElementEntity elemnt in _collection)
            {
                if (!elemnt.Name.ToLower().Contains(_searchName))
                    _showCollection.Remove(elemnt);
                else if (!_showCollection.Contains(elemnt))
                    _showCollection.Add(elemnt);
            }
        }

        /// <summary>
        /// Запуск без фильтрации
        /// </summary>
        [Obsolete("Данная кнопка не имеет смысла. Если нужен множественный выбор используй ElementMultiPick")]
        private void OnRunWithoutFilterClick(object sender, RoutedEventArgs e)
        {
            SelectedElement = null;
            _isRun = true;
            Status = RunStatus.Run;
            Close();
        }
    }
}
