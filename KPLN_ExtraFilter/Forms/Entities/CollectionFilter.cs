using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Data;

namespace KPLN_ExtraFilter.Forms.Entities
{
    /// <summary>
    /// Базовый фильтр для ICollectionView з поиском текстом
    /// </summary>
    public sealed class CollectionFilter<T> : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _searchText;
        private ICollectionView _view;

        /// <summary>
        /// Текст поиска для фильтра
        /// </summary>
        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value;
                NotifyPropertyChanged();
                _view?.Refresh();
            }
        }

        /// <summary>
        /// Гатовая ICollectionView для бинда в UI
        /// </summary>
        public ICollectionView View => _view;

        /// <summary>
        /// Инициализация нового ICollectionView
        /// </summary>
        public void LoadCollection(object sourceCollection)
        {
            if (sourceCollection == null)
            {
                _view = null;
                return;
            }

            _view = CollectionViewSource.GetDefaultView(sourceCollection);
            _view.Filter = FilterMethod;
            NotifyPropertyChanged(nameof(View));
        }

        /// <summary>
        /// Метод фильтрации
        /// </summary>
        public bool FilterMethod(object obj)
        {
            if (string.IsNullOrWhiteSpace(SearchText))
                return true;

            if (obj == null)
                return false;

            string text = string.Empty;
            if (obj is ParamEntity p)
                text = p.RevitParamName.ToLower();
            else if (obj is WSEntity ws)
                text = ws.RevitWSName.ToLower();
            else if (obj is CategoryEntity ce)
                text = ce.RevitCatName.ToLower();

            if (string.IsNullOrWhiteSpace(text))
                return false;

            string search = SearchText.ToLower();

            return text.Contains(search);
        }

        private void NotifyPropertyChanged([CallerMemberName] string name = "")
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
