using Autodesk.Revit.DB;
using KPLN_ExtraFilter.Common;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ExtraFilter.Forms.Entities
{
    public sealed class SelectionByModelM_ParamM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly SelectionByModelM _modelM;
        private Element[] _paramM_UserSelElems;
        private ParamEntity _paramM_SelectedCategory;

        public SelectionByModelM_ParamM(SelectionByModelM modelM)
        {
            _modelM = modelM;

            ParamM_Doc = modelM.Doc;
            ParamM_UserSelElems = modelM.Where_UserSelElems;

            ParamM_CategoryFilter.LoadCollection(ParamtM_DocCategories);
        }

        public Document ParamM_Doc { get; set; }

        public Element[] ParamM_UserSelElems
        {
            get => _paramM_UserSelElems;
            set
            {
                // запомним имя предыдущего выбора (если был)
                int prevParamId = _paramM_SelectedCategory == null ? -1 : _paramM_SelectedCategory.RevitParamIntId;

                _paramM_UserSelElems = value;
                NotifyPropertyChanged();

                ParamM_CategoryFilter.LoadCollection(ParamtM_DocCategories);
                ParamM_CategoryFilter.View?.Refresh();
                RestoreSelectedCategoryByName(prevParamId);

                NotifyPropertyChanged(nameof(ParamM_FilteredCategoryView));
            }
        }

        /// <summary>
        /// Экземпляр фильтра по категории
        /// </summary>
        public CollectionFilter<ParamEntity> ParamM_CategoryFilter { get; } = new CollectionFilter<ParamEntity>();

        /// <summary>
        /// Пользовательский ввод для поиска параметра
        /// </summary>
        public string SearchParamText
        {
            get => ParamM_CategoryFilter.SearchText;
            set => ParamM_CategoryFilter.SearchText = value;
        }

        public ICollectionView ParamM_FilteredCategoryView => ParamM_CategoryFilter.View;

        /// <summary>
        /// Коллекция параметров типа/экземпляра у выбранного эл-та
        /// </summary>
        public ParamEntity[] ParamtM_DocCategories
        {
            get
            {
                IEnumerable<Parameter> elemsParams = DocWorker.GetAllParamsFromElems(ParamM_Doc, ParamM_UserSelElems);
                if (elemsParams == null)
                    return null;

                List<ParamEntity> allParamsEntities = new List<ParamEntity>(elemsParams.Count());
                foreach (Parameter param in elemsParams)
                {
                    allParamsEntities.Add(new ParamEntity(param));
                }

                return allParamsEntities.OrderBy(pe => pe.RevitParamName).ToArray();
            }
        }

        /// <summary>
        /// Выбранный пользователем параметр
        /// </summary>
        public ParamEntity ParamM_SelectedParameter
        {
            get => _paramM_SelectedCategory;
            set
            {
                _paramM_SelectedCategory = value;
                NotifyPropertyChanged();
                _modelM.UpdateCanRunANDUserHelp();
            }
        }

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>
        /// Восстанавливает выбранный параметр, если в текущем View есть параметр с тем же id.
        /// </summary>
        private void RestoreSelectedCategoryByName(int prevParamId)
        {
            // если нет предыдущего выбора или View — просто ничего не делаем
            if (prevParamId == -1 || ParamM_CategoryFilter?.View == null)
                return;

            // Проходим по элементам View и ищем экземпляр с тем же именем
            var match = ParamM_CategoryFilter
                .View
                .Cast<ParamEntity>()
                .FirstOrDefault(c => c.RevitParamIntId == prevParamId);

            // присваиваем существующий экземпляр из коллекции — тогда ComboBox его увидит
            if (match != null)
                ParamM_SelectedParameter = match;
            // опционально: снять выбор, если совпадения нет
            else
                ParamM_SelectedParameter = null;
        }
    }
}
