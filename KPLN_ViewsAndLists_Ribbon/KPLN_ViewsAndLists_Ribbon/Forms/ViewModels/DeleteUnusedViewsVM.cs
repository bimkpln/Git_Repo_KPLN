using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_ViewsAndLists_Ribbon.ExecutableCommand;
using KPLN_ViewsAndLists_Ribbon.Forms.Commands;
using KPLN_ViewsAndLists_Ribbon.Forms.Entities;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KPLN_ViewsAndLists_Ribbon.Forms.MVVM
{
    public sealed class DeleteUnusedViewsVM
    {
        private readonly DeleteUnusedViewsForm _mainWindow;

        public DeleteUnusedViewsVM(DeleteUnusedViewsForm mainWindow, UIApplication uiapp, View[] docViews)
        {
            _mainWindow = mainWindow;
            
            CurrentDeleteUnusedViewsM = new DeleteUnusedViewsM(_mainWindow, uiapp, docViews);

            DeleteSelectedCmd = new RelayCommand<object>(DeleteSelected);
            SelectAllCmd = new RelayCommand<object>(_ => SelectAll());
            DeselectAllCmd = new RelayCommand<object>(_ => DeselectAll());
            ReverseCmd = new RelayCommand<object>(_ => Reverse());
            CloseWindowCmd = new RelayCommand<object>(CloseWindow);
        }

        public DeleteUnusedViewsM CurrentDeleteUnusedViewsM { get; set; }

        /// <summary>
        /// Комманда: Удалить выбранные
        /// </summary>
        public ICommand DeleteSelectedCmd { get; }

        /// <summary>
        /// Комманда: Выбрать все
        /// </summary>
        public ICommand SelectAllCmd { get; }

        /// <summary>
        /// Комманда: Снять со всех
        /// </summary>
        public ICommand DeselectAllCmd { get; }

        /// <summary>Реверс выбора 
        /// </summary>
        public ICommand ReverseCmd { get; }

        /// <summary>
        /// Комманда: Закрыть окно
        /// </summary>
        public ICommand CloseWindowCmd { get; }

        public void DeleteSelected(object windObj)
        {
            CloseWindow(windObj);
            KPLN_Loader.Application.OnIdling_CommandQueue.Enqueue(new DeleteElemsExcCmd(CurrentDeleteUnusedViewsM));
        }
        
        public void SelectAll()
        {
            foreach (TreeElementEntity root in CurrentDeleteUnusedViewsM.TreeElemEntities)
            {
                // Включаю корни (ветки+листья - через сеттер автоматом)
                root.IsChecked = true;
            }
        }

        public void DeselectAll()
        {
            List<TreeElementEntity> leafElems = new List<TreeElementEntity>();
            List<TreeElementEntity> branchElems = new List<TreeElementEntity>();

            foreach (TreeElementEntity root in CurrentDeleteUnusedViewsM.TreeElemEntities)
            {
                // Отключаю корни
                root.IsChecked = false;

                GetTEEElems(root, leafElems);
                GetTEEBranches(root, branchElems);
            }

            // Отключаю только листья
            foreach (TreeElementEntity leaf in leafElems)
            {
                leaf.IsChecked = false;
            }

            // Отключаю все ветки
            foreach (TreeElementEntity branch in branchElems)
            {
                branch.IsChecked = false;
            }
        }

        public void Reverse()
        {
            List<TreeElementEntity> leafElems = new List<TreeElementEntity>();
            List<TreeElementEntity> branchElems = new List<TreeElementEntity>();

            foreach (TreeElementEntity root in CurrentDeleteUnusedViewsM.TreeElemEntities)
            {
                GetTEEElems(root, leafElems);
                GetTEEBranches(root, branchElems);
            }

            // Реверсирую только листья
            foreach (TreeElementEntity leaf in leafElems)
            {
                leaf.IsChecked = !leaf.IsChecked;
            }

            // Запоминаем итоговое состояние листьев после реверса
            Dictionary<TreeElementEntity, bool> leafStates = leafElems
                .ToDictionary(leaf => leaf, leaf => leaf.IsChecked);

            // Реверсирую все ветки
            foreach (TreeElementEntity branch in branchElems)
            {
                branch.IsChecked = branch.TEE_ChildrenColl.All(child => child.IsChecked);
            }

            // ВАЖНО:
            // если branch.IsChecked = false, setter может снять галки с детей.
            // Поэтому возвращаем листьям их правильное состояние.
            foreach (KeyValuePair<TreeElementEntity, bool> pair in leafStates)
            {
                pair.Key.IsChecked = pair.Value;
            }

            // И финально ещё раз пересчитываем ветки после восстановления листьев
            foreach (TreeElementEntity branch in branchElems)
            {
                branch.IsChecked = branch.TEE_ChildrenColl.All(child => child.IsChecked);
            }
        }

        public void CloseWindow(object windObj)
        {
            if (windObj is Window window)
                window.Close();
        }

        /// <summary>
        /// Получить коллекцию веток и корней.
        /// Порядок: сначала нижние ветки, потом верхние.
        /// </summary>
        private static void GetTEEBranches(TreeElementEntity root, List<TreeElementEntity> result)
        {
            if (root == null)
                return;

            if (root.TEE_ChildrenColl == null || root.TEE_ChildrenColl.Count == 0)
                return;

            foreach (TreeElementEntity child in root.TEE_ChildrenColl)
            {
                GetTEEBranches(child, result);
            }

            result.Add(root);
        }

        /// <summary>
        /// Получить коллекцию конечных элементов
        /// </summary>
        private static void GetTEEElems(TreeElementEntity root, List<TreeElementEntity> result)
        {
            if (root == null)
                return;

            if (root.TEE_Element != null)
                result.Add(root);

            if (root.TEE_ChildrenColl == null || root.TEE_ChildrenColl.Count == 0)
                return;

            foreach (TreeElementEntity child in root.TEE_ChildrenColl)
            {
                GetTEEElems(child, result);
            }
        }
    }
}
