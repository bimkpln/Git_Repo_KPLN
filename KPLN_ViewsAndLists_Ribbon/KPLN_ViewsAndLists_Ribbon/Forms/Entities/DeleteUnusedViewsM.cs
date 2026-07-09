using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace KPLN_ViewsAndLists_Ribbon.Forms.Entities
{
    public sealed class DeleteUnusedViewsM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private TreeElementEntity[] _treeElemEntities;
        private string _searchText;

        public DeleteUnusedViewsM(DeleteUnusedViewsForm mainWindow, UIApplication uiapp, IEnumerable<View> docViews)
        {
            MainWindow = mainWindow;
            UIApp = uiapp;
            DocViews = docViews.ToArray();
            
            CreateTree(DocViews);
        }

        public DeleteUnusedViewsForm MainWindow { get; }

        public UIApplication UIApp { get; }

        /// <summary>
        /// Коллекция видов модели
        /// </summary>
        public View[] DocViews { get; }

        /// <summary>
        /// Пользовательский ввод для поиска
        /// </summary>
        public string SearchText
        {
            get => _searchText;
            set
            {
                _searchText = value;
                NotifyPropertyChanged();

                ApplyTreeFilter();
            }
        }

        /// <summary>
        /// Коллекция элементов в дереве
        /// </summary>
        public TreeElementEntity[] TreeElemEntities
        {
            get => _treeElemEntities;
            set
            {
                _treeElemEntities = value;
                NotifyPropertyChanged();
            }
        }

        private void ApplyTreeFilter()
        {
            if (TreeElemEntities == null)
                return;

            foreach (TreeElementEntity rootEntity in TreeElemEntities)
            {
                ApplyTreeFilter(rootEntity, SearchText);
            }
        }

        private bool ApplyTreeFilter(TreeElementEntity entity, string searchText)
        {
            bool isSearchEmpty = string.IsNullOrWhiteSpace(searchText);

            bool selfContains = isSearchEmpty ||
                entity.TEE_Name
                    .ToLower()
                    .Contains(searchText.ToLower());

            bool childContains = false;

            if (entity.TEE_ChildrenColl != null)
            {
                foreach (TreeElementEntity child in entity.TEE_ChildrenColl)
                {
                    if (ApplyTreeFilter(child, searchText))
                        childContains = true;
                }
            }

            bool isVisible = selfContains || childContains;
            entity.IsVisible = isVisible;

            // Опционально: раскрывать ветки, где есть совпадение
            if (!isSearchEmpty && childContains)
                entity.IsExpanded = true;

            return isVisible;
        }

        /// <summary>
        /// Сгенерировать дерево в окне
        /// </summary>
        private void CreateTree(IEnumerable<Element> userSelElems) =>
            TreeElemEntities = TreeElementEntity.CreateTreeElEnt_ByCatANDFam(userSelElems);

        private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
