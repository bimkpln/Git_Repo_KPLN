using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using KPLN_ViewsAndLists_Ribbon.Forms.MVVM;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;


namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    public partial class DeleteUnusedViewsForm : Window
    {
        private TreeElementEntity _lastClickedEntity;

        public DeleteUnusedViewsForm(UIApplication uiapp, View[] docViews)
        {
            CurrentDeleteUnusedViewsVM = new DeleteUnusedViewsVM(this, uiapp, docViews);

            InitializeComponent();

            this.tbUserInput.Focus();
            DataContext = CurrentDeleteUnusedViewsVM;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public DeleteUnusedViewsVM CurrentDeleteUnusedViewsVM { get; set; }

        /// <summary>
        /// Отлов клика по элементу дерева, для добавления управления Shift'ом
        /// </summary>
        private void TreeElementCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is CheckBox checkBox) || !(checkBox.DataContext is TreeElementEntity currentEntity))
                return;

            bool isChecked = checkBox.IsChecked == true;

            // Метка использования шифта. Если без неё - просто помечаем предыдущий клик
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == 0 || _lastClickedEntity == null)
            {
                _lastClickedEntity = currentEntity;
                return;
            }

            // Shift + клик: ставим галки между _lastClickedItem и currentEntity
            if (!(DataContext is DeleteUnusedViewsVM vm) || vm.CurrentDeleteUnusedViewsM?.TreeElemEntities == null)
            {
                _lastClickedEntity = currentEntity;
                return;
            }

            var flatList = new List<TreeElementEntity>();
            foreach (var root in vm.CurrentDeleteUnusedViewsM.TreeElemEntities)
                FlattenTree(root, flatList);

            int index1 = flatList.IndexOf(_lastClickedEntity);
            int index2 = flatList.IndexOf(currentEntity);

            if (index1 == -1 || index2 == -1)
            {
                _lastClickedEntity = currentEntity;
                return;
            }

            if (index2 < index1)
            {
                int tmp = index1;
                index1 = index2;
                index2 = tmp;
            }

            for (int i = index1; i <= index2; i++)
                flatList[i].IsChecked = isChecked;

            _lastClickedEntity = currentEntity;
        }

        /// <summary>
        /// Выпрямленный список ВСЕХ элементов дерева
        /// </summary>
        private void FlattenTree(TreeElementEntity node, List<TreeElementEntity> result)
        {
            result.Add(node);

            if (node.TEE_ChildrenColl == null)
                return;

            foreach (var child in node.TEE_ChildrenColl)
                FlattenTree(child, result);
        }
    }
}