using KPLN_Library_Forms.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class ElementSinglePick : Window
    {
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

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(SearchText);

        private void OnElementClick(object sender, RoutedEventArgs e)
        {
            SelectedElement = (sender as Button).DataContext as ElementEntity;
            this.DialogResult = true;

            Close();
        }

        /// <summary>
        /// Фильтрация по имени
        /// </summary>
        private void SearchText_Changed(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string _searchName = textBox.Text.ToLower();

            _showCollection.Clear();
            foreach (ElementEntity elemnt in _collection)
            {
                if (elemnt.Name.ToLower().Contains(_searchName))
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
            this.DialogResult = true;

            Close();
        }
    }
}
