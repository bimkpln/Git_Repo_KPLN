using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_ExtraFilter.ExternalEventHandler;
using KPLN_ExtraFilter.Forms.Entities;
using KPLN_ExtraFilter.Forms.ViewModels;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_ExtraFilter.Forms
{
    public partial class SelectionByModel : Window
    {
        private static SelectionByModel _currentInstance;

        private readonly bool _isUpdateble;
        private readonly ExternalEvent _subscriptionExtEv;
        private SelectionByModelSubscriptionHandler _subscriptionHandler;

        private readonly ExternalEvent _viewExtEv;
        private ViewActivatedHandler _viewHandler;

#if !Debug2020 && !Revit2020
        private readonly ExternalEvent _selExtEv;
        private SelectionChangedHandler _selHandler;
#endif

        private TreeElementEntity _lastClickedEntity;

        public SelectionByModel(UIApplication uiapp, ViewFilterMode viewFilterMode, bool isUpdateble)
        {
            _isUpdateble = isUpdateble;

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
                _viewExtEv = FormEventSubscriptionHelper.CreateViewChangedEvent(handler =>
                {
                    _viewHandler = handler;
                    _viewHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;
                });

#if !Debug2020 && !Revit2020
                _selExtEv = FormEventSubscriptionHelper.CreateSelectionChangedEvent(handler =>
                {
                    _selHandler = handler;
                    _selHandler.CurrentSelByModelVM = CurrentSelectionByModelVM;
                });
#endif

                _subscriptionHandler = new SelectionByModelSubscriptionHandler()
                {
                    ViewHandler = OnViewChanged,
#if !Debug2020 && !Revit2020
                    SelectionHandler = OnSelectionChanged,
#endif
                };
                _subscriptionExtEv = ExternalEvent.Create(_subscriptionHandler);

                CurrentSelectionByModelVM.CurrentSelectionByModelM.PropertyChanged += OnModelPropertyChanged;
                UpdateEventSubscriptions();
            }
            #endregion

            _currentInstance = this;
            Closed += (sender, args) =>
            {
                if (isUpdateble)
                {
                    CurrentSelectionByModelVM.CurrentSelectionByModelM.PropertyChanged -= OnModelPropertyChanged;
                    RequestEventSubscriptions(false, false);
                }

                if (ReferenceEquals(_currentInstance, this))
                    _currentInstance = null;
            };
        }

        /// <summary>
        /// Активировать уже открытое окно
        /// </summary>
        /// <returns>True, если открытое окно найдено</returns>
        public static bool TryActivateExisting()
        {
            if (_currentInstance == null || !_currentInstance.IsLoaded)
            {
                _currentInstance = null;
                return false;
            }

            if (_currentInstance.WindowState == WindowState.Minimized)
                _currentInstance.WindowState = WindowState.Normal;

            if (!_currentInstance.IsVisible)
                _currentInstance.Show();

            _currentInstance.Activate();
            _currentInstance.Focus();

            return true;
        }

        /// <summary>
        /// VM для окна
        /// </summary>
        public SelectionByModelVM CurrentSelectionByModelVM { get; set; }

        private void OnModelPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(SelectionByModelM.Where_ViewDocFilterMode))
                UpdateEventSubscriptions();
        }

        private void UpdateEventSubscriptions()
        {
            if (!_isUpdateble || CurrentSelectionByModelVM?.CurrentSelectionByModelM == null)
                return;

            ViewFilterMode mode = CurrentSelectionByModelVM.CurrentSelectionByModelM.Where_ViewDocFilterMode;
            RequestEventSubscriptions(
                mode == ViewFilterMode.CurrentView,
                mode == ViewFilterMode.UserSelection);
        }

        private void RequestEventSubscriptions(bool subscribeViewChanged, bool subscribeSelectionChanged)
        {
            if (_subscriptionHandler == null || _subscriptionExtEv == null)
                return;

            _subscriptionHandler.ShouldSubscribeViewChanged = subscribeViewChanged;

#if !Debug2020 && !Revit2020
            _subscriptionHandler.ShouldSubscribeSelectionChanged = subscribeSelectionChanged;
#endif

            _subscriptionExtEv.Raise();
        }

        private void OnViewChanged(object sender, ViewActivatedEventArgs e) => _viewExtEv?.Raise();

#if !Debug2020 && !Revit2020
        private void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e) => _selExtEv?.Raise();
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