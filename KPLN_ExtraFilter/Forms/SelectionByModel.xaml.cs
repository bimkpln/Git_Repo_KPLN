using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByModel : Window
    {
        private ExternalEvent _viewExtEv;
        private ViewActivatedHandler _viewHandler;

        private ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;

        private TreeElementEntity _lastClickedEntity;

        public SelectionByModel(UIApplication uiapp, ViewFilterMode viewFilterMode, bool isUpdateble)
        {
            CurrentSelectionByModelVM = new SelectionByModelVM(this, uiapp, viewFilterMode, isUpdateble);

            InitializeComponent();

            DataContext = CurrentSelectionByModelVM;

#if Debug2020 || Revit2020
            // Нет метода в API для отслеживания изменний в выборке юзера
            this.AddToCurrent.Visibility = System.Windows.Visibility.Hidden;
#endif

#region Блок настройки подписки на события окна на происходящие триггеры
            if (isUpdateble)
            {
                // Создание ExternalEvent для переключения видов
                ViewActivatedHandler viewHandler = new ViewActivatedHandler();
                ExternalEvent viewExtEv = ExternalEvent.Create(viewHandler);

                // Создание ExternalEvent для отписки от переключения видов (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
                UnsubViewActHandler unsubViewHandler = new UnsubViewActHandler() { Handler = OnViewChanged };
                ExternalEvent unsubViewExtEv = ExternalEvent.Create(unsubViewHandler);

    #if !Debug2020 && !Revit2020
                // Создание ExternalEvent для выделения эл-в
                SelectionChangedHandler selHandler = new SelectionChangedHandler();
                ExternalEvent selExtEv = ExternalEvent.Create(selHandler);

                // Создание ExternalEvent для отписки от выбора эл-в (конструкция для возвращения в контекст ревит, т.к. wpf из него выпадает)
                UnsubEventHandler unsubSelHandler = new UnsubEventHandler() { Handler = OnSelectionChanged };
                ExternalEvent unsubSelExtEv = ExternalEvent.Create(unsubSelHandler);

                // Доп настройки окна
                this.SetExternalEvent(viewExtEv, viewHandler, selExtEv, selHandler);
    #endif

                // Подписываюсь на OnViewChanged
                uiapp.ViewActivated += OnViewChanged;
                // Подписываю окно на отписку (через ExternalEvent)
                this.Closed += (s, e) => unsubViewExtEv.Raise();

    #if !Debug2020 && !Revit2020
                // Подписываюсь на SelectionChanged
                uiapp.SelectionChanged += OnSelectionChanged;
                // Подписываю окно на отписку (через ExternalEvent)
                this.Closed += (s, e) => unsubSelExtEv.Raise();
    #endif
            }
#endregion
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByModelVM CurrentSelectionByModelVM { get; set; }

        public void SetExternalEvent(ExternalEvent viewExtEv, ViewActivatedHandler viewHandler, ExternalEvent selExtEv, SelectionChangedHandler selHandler)
        {
            _viewExtEv = viewExtEv;
            
            _viewHandler = viewHandler;
            _viewHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;


            _selExtEv = selExtEv;
            
            _selHandler = selHandler;
            _selHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;
        }

        public void RaiseUpdateSelChanged() => _selExtEv?.Raise();

        public void RaiseUpdateViewChanged() => _viewExtEv?.Raise();

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => this.RaiseUpdateViewChanged();

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => this.RaiseUpdateSelChanged();
#endif

        private void CHB_Where_Workset_Checked(object sender, RoutedEventArgs e) => this.CB_FilterWS.Focus();

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
            if (!(DataContext is SelectionByModelVM vm) || vm.CurrentSelectionByModelM?.TreeElemEntities == null)
            {
                _lastClickedEntity = currentEntity;
                return;
            }

            var flatList = new List<TreeElementEntity>();
            foreach (var root in vm.CurrentSelectionByModelM.TreeElemEntities)
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
