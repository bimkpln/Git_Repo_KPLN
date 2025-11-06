using Autodesk.Revit.UI;
using KPLN_Library_Forms.Services;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Common;
using KPLN_TaskManager.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace KPLN_TaskManager.Forms
{
    public partial class TaskManagerView : Page, IDockablePaneProvider, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<TaskItemEntity> _collection;

        // Подотделы для отдела пользователя
        private readonly DBSubDepartment[] _subDepsForUser;
        private string _searchHeader;
        private string _selectedOpenStausTasks = "Open";
        private bool _fullTaskCollection = false;
        private string _subDepDependence = "AllSuDepTask";

        public TaskManagerView()
        {
            _subDepsForUser = DBMainService.DBSubDepartmentColl
                .Where(sd => sd.Id == DBMainService.CurrentUserDBSubDepartment.Id || sd.DependentSubDepId == DBMainService.CurrentUserDBSubDepartment.Id)
                .ToArray();

            InitializeComponent();
            
            DataContext = this;
        }

        public ICollectionView FilteredTasks { get; set; }

        /// <summary>
        /// Фильтрация по заголовку
        /// </summary>
        public string SearchHeader
        {
            get => _searchHeader;
            set
            {
                _searchHeader = value;
                OnPropertyChanged();

                FilteredTasks?.Refresh();
            }
        }

        /// <summary>
        /// Фильтрация по открым замечаниям
        /// </summary>
        public string SelectedOpenStausTasks
        {
            get => _selectedOpenStausTasks;
            set
            {
                _selectedOpenStausTasks = value;
                OnPropertyChanged();

                FilteredTasks?.Refresh();
            }
        }

        public bool FullTaskCollection
        {
            get => _fullTaskCollection;
            set
            {
                _fullTaskCollection = value;
                OnPropertyChanged();

                FilteredTasks?.Refresh();
                UpdateSubDepSelectionEnabled();
            }
        }

        /// <summary>
        /// Фильтрация по адресату задания (задание входящее, задание исходящее и т.п.)
        /// </summary>
        public string SubDepDependence
        {
            get => _subDepDependence;
            set
            {
                _subDepDependence = value;
                OnPropertyChanged();

                FilteredTasks?.Refresh();
            }
        }


        /// <summary>
        /// Загрузить таски в окно
        /// </summary>
        /// <returns></returns>
        public TaskManagerView LoadTaskData()
        {
            if (Module.CurrentDBProject == null)
                return null;

            _collection = new ObservableCollection<TaskItemEntity>(TMDBService.GetEntities_ByDBProject(Module.CurrentDBProject));

            FilteredTasks = CollectionViewSource.GetDefaultView(_collection);
            FilteredTasks.Filter = FilterTasks;

            // Оповещаю о изменении коллекции
            OnPropertyChanged(nameof(FilteredTasks));

            return this;
        }

        /// <summary>
        /// Метод фильтрации элементов
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool FilterTasks(object item)
        {
            if (item is TaskItemEntity task)
            {
                bool matchesTitle = string.IsNullOrWhiteSpace(SearchHeader) || task.TaskHeader.ToLower().Contains(SearchHeader.ToLower());

                bool matchesStatus = SelectedOpenStausTasks == "All" || SelectedOpenStausTasks == task.TaskStatus.ToString();

                if (FullTaskCollection)
                    return matchesTitle && matchesStatus;

                bool userProp = GetUserPropValue(task);

                // Поиск по отделу и подотделу
                bool isInputSubDepTask = DBMainService.CurrentUserDBSubDepartment.Id == task.DelegatedDepartmentId
                    || _subDepsForUser.Any(sdu => task.DelegatedDepartmentId == sdu.Id);
                bool isOutputSubDepTask = DBMainService.CurrentUserDBSubDepartment.Id == task.CreatedTaskDepartmentId 
                    || _subDepsForUser.Any(sdu => task.CreatedTaskDepartmentId == sdu.Id);

                if (SubDepDependence == "AllSuDepTask")
                    return matchesTitle && matchesStatus && userProp && (isInputSubDepTask || isOutputSubDepTask);

                if (SubDepDependence == "InputSubDepTask")
                    return matchesTitle && matchesStatus && isInputSubDepTask;

                if (SubDepDependence == "OutputSubDepTask")
                    return matchesTitle && matchesStatus && isOutputSubDepTask;

                string openDocTask = "Nope";
                if (task.ModelName != null && !string.IsNullOrEmpty(task.ModelName))
                    openDocTask = Module.CurrentFileName.Contains(task.ModelName) ? "OpenDocTask" : "Nope";
                bool isOpenDocTask = SubDepDependence.Equals(openDocTask);
                if (SubDepDependence == "OpenDocTask")
                    return matchesTitle && matchesStatus && isOpenDocTask;

                return true;
            }
            return false;
        }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = this;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Left
            };
        }

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private void TaskItem_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.DataContext is TaskItemEntity taskItemEntity)
            {
                // Открываю окно
                TaskItemView taskItemView = new TaskItemView(taskItemEntity);
                WindowHandleSearch.MainWindowHandle.SetAsOwner(taskItemView);

                // Слежу за изменением статуса сущности в открываемом окне
                taskItemView.CurrentTaskItemEntity.PropertyChanged += (s, args) =>
                {
                    if (args.PropertyName == nameof(TaskItemEntity.TaskStatus))
                        FilteredTasks?.Refresh();
                };

                // Слежу за закрытием окна
                taskItemView.Closed += (s, args) => LoadTaskData();

                taskItemView.Show();
            }
        }

        private void AddTask_Click(object sender, RoutedEventArgs e)
        {
            if (Module.CurrentDBProject == null)
            {
                MessageBox.Show(
                    $"Работает только с проектами КПЛН, которые хранятся по соответсвующей структуре.\n\nПримечание: На отсоединенных файлах тоже НЕ работает",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                return;
            }

            TaskItemEntity taskItemEntity = new TaskItemEntity(Module.CurrentDBProject.Id, DBMainService.CurrentDBUser.Id, DBMainService.CurrentUserDBSubDepartment.Id);

            // Открываю окно
            TaskItemView taskItemView = new TaskItemView(taskItemEntity);

            WindowHandleSearch.MainWindowHandle.SetAsOwner(taskItemView);
            taskItemView.Closed += (s, args) => LoadTaskData();
            taskItemView.Show();
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            string path = ExportToExcelService.SetPath();

            if (!string.IsNullOrEmpty(path))
            {
                List<TaskItemEntity> filteredTasks = new List<TaskItemEntity>();
                if (sender is MenuItem mItem)
                {
                    TaskItemEntity task = mItem.DataContext as TaskItemEntity;
                    filteredTasks.Add(task);
                }
                else
                {
                    foreach (TaskItemEntity task in FilteredTasks)
                    {
                        filteredTasks.Add(task);
                    }
                }

                ExportToExcelService.Run(path, Module.CurrentDBProject.Name, filteredTasks);
            }
        }

        private void UpdateData_Click(object sender, RoutedEventArgs e) => LoadTaskData();

        /// <summary>
        /// Обновить управляемость RadioButton группы "SubDepSelection"
        /// </summary>
        private void UpdateSubDepSelectionEnabled()
        {
            if (FullTaskCollection)
            {
                AllTaskRBtn.IsChecked = true;
                AllTaskRBtn.IsEnabled = false;
                InputTaskRBtn.IsEnabled = false;
                OutputTaskRBtn.IsEnabled = false;
                OpenDocTaskRBtn.IsEnabled = false;
            }
            else
            {
                AllTaskRBtn.IsEnabled = true;
                InputTaskRBtn.IsEnabled = true;
                OutputTaskRBtn.IsEnabled = true;
                OpenDocTaskRBtn.IsEnabled = true;
            }
        }

        /// <summary>
        /// Кастомная настройка видимости TaskItemEntity для конкретного пользователя
        /// </summary>
        /// <param name="task"></param>
        /// <returns></returns>
        private bool GetUserPropValue(TaskItemEntity task)
        {
            //Для бим-отдела верстка на лету под спецов, чтобы они видели замечания по своим отделам
            if (DBMainService.CurrentUserDBSubDepartment.Id == 8)
            {
                switch (DBMainService.CurrentDBUser.Surname)
                {
                    case "Куцко":
                        return true;
                    case "Коломиец":
                        goto case "Федосеева";
                    case "Федосеева":
                        return task.DelegatedDepartmentId == 2 
                            || task.CreatedTaskDepartmentId == 2 
                            || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
                    case "Тарчоков":
                        goto case "Ямковой";
                    case "Ямковой":
                        return task.DelegatedDepartmentId == 3 
                            || task.CreatedTaskDepartmentId == 3
                            || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
                    case "Садовская":
                        return task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                            || task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                            || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21
                            || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22
                            || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
                    case "Чичева":
                        return task.DelegatedDepartmentId == 4 || task.CreatedTaskDepartmentId == 4
                            || task.DelegatedDepartmentId == 6 || task.CreatedTaskDepartmentId == 6
                            || task.DelegatedDepartmentId == 20 || task.CreatedTaskDepartmentId == 20
                            || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
                }
            }

            //Для руководителей - верстка на лету под их отделы, чтобы они видели замечания по своим отделам
            if (DBMainService.CurrentDBUser.Surname == "Кудрова")
                return task.DelegatedDepartmentId == 4 || task.CreatedTaskDepartmentId == 4
                    || task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                    || task.DelegatedDepartmentId == 20 || task.CreatedTaskDepartmentId == 20
                    || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21
                    || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
            else if (DBMainService.CurrentDBUser.Surname == "Тамарин")
                return task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                    || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21
                    || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
            else if (DBMainService.CurrentDBUser.Surname == "Колодий")
                return task.DelegatedDepartmentId == 6 || task.CreatedTaskDepartmentId == 6
                    || task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                    || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22
                    || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;
            else if (DBMainService.CurrentDBUser.Surname == "Алиев")
                return task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                    || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22
                    || task.CreatedTaskUserId == DBMainService.CurrentDBUser.Id;

            return true;
        }
    }
}
