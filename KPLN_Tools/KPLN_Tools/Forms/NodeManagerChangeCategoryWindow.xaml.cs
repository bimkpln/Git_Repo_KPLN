using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace KPLN_Tools.Forms
{
    public partial class ChangeCategoryWindowNodeManager : Window
    {
        public class CategoryItem
        {
            public int CatId { get; set; }
            public string Name { get; set; }
            public List<SubcatItem> Subcats { get; set; } = new List<SubcatItem>();
            public override string ToString() => Name;
        }

        public class SubcatItem
        {
            public string Id { get; set; }     
            public string Title { get; set; }  
            public override string ToString() => Title;
        }

        private readonly List<CategoryItem> _categories;

        public int? SelectedCatId { get; private set; }
        public string SelectedSubcatId { get; private set; }

        public ChangeCategoryWindowNodeManager(
            IEnumerable<NodeCategoryUi> categoryTree,
            int? currentCatId,
            string currentSubcatId,
            string elementName)
        {
            InitializeComponent();

            Title = elementName; 

            _categories = BuildCategoryItems(categoryTree);
            CbCategory.ItemsSource = _categories;

            if (currentCatId.HasValue)
            {
                var cat = _categories.FirstOrDefault(c => c.CatId == currentCatId.Value);
                if (cat != null)
                {
                    CbCategory.SelectedItem = cat;
                    PopulateSubcats(cat, currentSubcatId);
                }
                else if (_categories.Count > 0)
                {
                    CbCategory.SelectedIndex = 0;
                    PopulateSubcats(_categories[0], null);
                }
            }
            else if (_categories.Count > 0)
            {
                CbCategory.SelectedIndex = 0;
                PopulateSubcats(_categories[0], null);
            }
        }

        private List<CategoryItem> BuildCategoryItems(IEnumerable<NodeCategoryUi> categoryTree)
        {
            var list = new List<CategoryItem>();

            foreach (var root in categoryTree.Where(r => r.Parent == null))
            {
                var catItem = new CategoryItem
                {
                    CatId = root.CatId,
                    Name = root.Title
                };

                var subList = new List<SubcatItem>();

                foreach (var child in root.Children)
                {
                    FlattenSubcats(child, subList, child.Title);
                }

                if (subList.Count == 0)
                {
                    subList.Add(new SubcatItem
                    {
                        Id = "0",
                        Title = "0 - Без категории"
                    });
                }

                catItem.Subcats = subList;
                list.Add(catItem);
            }

            return list;
        }

        private void FlattenSubcats(NodeCategoryUi node, List<SubcatItem> list, string path)
        {
            list.Add(new SubcatItem
            {
                Id = node.Id,
                Title = $"{node.Id} - {path}"
            });

            foreach (var child in node.Children)
            {
                FlattenSubcats(child, list, path + " > " + child.Title);
            }
        }

        private void PopulateSubcats(CategoryItem cat, string currentSubcatId)
        {
            CbSubcategory.ItemsSource = cat.Subcats;

            if (string.IsNullOrEmpty(currentSubcatId))
            {
                if (cat.Subcats.Count > 0)
                    CbSubcategory.SelectedIndex = 0;
            }
            else
            {
                var sub = cat.Subcats.FirstOrDefault(s => s.Id == currentSubcatId);
                if (sub != null)
                    CbSubcategory.SelectedItem = sub;
                else if (cat.Subcats.Count > 0)
                    CbSubcategory.SelectedIndex = 0;
            }
        }

        private void CbCategory_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cat = CbCategory.SelectedItem as CategoryItem;
            if (cat == null)
            {
                CbSubcategory.ItemsSource = null;
                return;
            }

            PopulateSubcats(cat, null);
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var cat = CbCategory.SelectedItem as CategoryItem;
            var sub = CbSubcategory.SelectedItem as SubcatItem;

            if (cat == null)
            {
                MessageBox.Show("Выберите категорию.", "KPLN. Менеджер узлов",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (sub == null)
            {
                MessageBox.Show("Выберите подкатегорию.", "KPLN. Менеджер узлов",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            SelectedCatId = cat.CatId;
            SelectedSubcatId = sub.Id;

            DialogResult = true;
            Close();
        }
    }
}
