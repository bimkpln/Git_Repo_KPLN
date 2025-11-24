using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    public sealed class SelectionByModelM_CategoryM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly SelectionByModelM _modelM;
        private Element[] _catM_UserSelElems;
        private CategoryEntity _catM_SelectedCategory;

        public SelectionByModelM_CategoryM(SelectionByModelM modelM)
        {
            _modelM = modelM;
            
            CatM_Doc = modelM.Doc;
            CatM_UserSelElems = modelM.Cahce_UserSelElemsWithoutCatFilter;

            CatM_CategoryFilter.LoadCollection(CatM_DocCategories);
        }

        public Document CatM_Doc { get; set; }

        public Element[] CatM_UserSelElems
        {
            get => _catM_UserSelElems;
            set
            {
                // запомним имя предыдущего выбора (если был)
                string prevName = _catM_SelectedCategory?.RevitCatName;

                _catM_UserSelElems = value;
                NotifyPropertyChanged();
                
                CatM_CategoryFilter.LoadCollection(CatM_DocCategories);
                CatM_CategoryFilter.View?.Refresh();
                RestoreSelectedCategoryByName(prevName);

                NotifyPropertyChanged(nameof(CatM_FilteredCategoryView));
            }
        }

        /// <summary>
        /// Экземпляр фильтра по категории
        /// </summary>
        public CollectionFilter<CategoryEntity> CatM_CategoryFilter { get; } = new CollectionFilter<CategoryEntity>();

        /// <summary>
        /// Пользовательский ввод для поиска параметра
        /// </summary>
        public string SearchCategoryText
        {
            get => CatM_CategoryFilter.SearchText;
            set => CatM_CategoryFilter.SearchText = value;
        }

        public ICollectionView CatM_FilteredCategoryView => CatM_CategoryFilter.View;

        /// <summary>
        /// Коллекция категорий модели
        /// </summary>
        public CategoryEntity[] CatM_DocCategories
        {
            get
            {
                if (CatM_Doc == null || CatM_UserSelElems == null || CatM_UserSelElems.Count() == 0)
                    return null;

                return DocWorker.GetAllCatsFromElems(CatM_UserSelElems).Select(cat => new CategoryEntity(cat)).OrderBy(cat => cat.RevitCatName).ToArray();
            }
        }

        /// <summary>
        /// Выбранная пользователем категория
        /// </summary>
        public CategoryEntity CatM_SelectedCategory
        {
            get => _catM_SelectedCategory;
            set
            {
                _catM_SelectedCategory = value;
                NotifyPropertyChanged();
                _modelM.UpdateCanRunANDUserHelp();
                _modelM.SetUserSelElems();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>
        /// Восстанавливает выбранную категорию, если в текущем View есть категория с тем же именем.
        /// </summary>
        private void RestoreSelectedCategoryByName(string prevRevitCatName)
        {
            // если нет предыдущего выбора или View — просто ничего не делаем
            if (string.IsNullOrWhiteSpace(prevRevitCatName) || CatM_CategoryFilter?.View == null)
                return;

            // Проходим по элементам View и ищем экземпляр с тем же именем
            var match = CatM_CategoryFilter
                .View
                .Cast<CategoryEntity>()
                .FirstOrDefault(c => c.RevitCatName == prevRevitCatName);

            // присваиваем существующий экземпляр из коллекции — тогда ComboBox его увидит
            if (match != null)
                CatM_SelectedCategory = match;
            // опционально: снять выбор, если совпадения нет
            else
                CatM_SelectedCategory = null;
        }
    }
}
