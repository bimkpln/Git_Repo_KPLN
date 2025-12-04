using KPLN_Library_Forms.Common;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Library_Forms.UI
{
    public partial class ElementMultiPick : Window
    {
        private readonly ObservableCollection<ElementEntity> _collection;
        private readonly ObservableCollection<ElementEntity> _showCollection;

        public ElementMultiPick(Window owner, IEnumerable<ElementEntity> collection)
        {
            _collection = new ObservableCollection<ElementEntity>(collection);
            _showCollection = new ObservableCollection<ElementEntity>(_collection);
            InitializeComponent();

            this.Title = $"KPLN: Выбери нужный элемент/ы";
            this.Owner = owner;
            Elements.ItemsSource = _showCollection;
            PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        public ElementMultiPick(Window owner, IEnumerable<ElementEntity> collection, string title) : this(owner, collection)
        {
            this.Title = $"KPLN: {title}";
        }

        /// <summary>
        /// Выбранная коллекция элементов
        /// </summary>
        public List<ElementEntity> SelectedElements => _collection.Where(ee => ee.IsSelected).ToList();

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e) => Keyboard.Focus(SearchText);

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

        private void RunBtn_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CheckAllBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (object obj in Elements.Items)
            {
                ElementEntity entity = (ElementEntity)obj;
                entity.IsSelected = true;
            }
        }

        private void ToggleBtn_Click(object sender, RoutedEventArgs e)
        {
            foreach (object obj in Elements.Items)
            {
                ElementEntity entity = (ElementEntity)obj;
                entity.IsSelected = !entity.IsSelected;
            }
        }
    }
}
