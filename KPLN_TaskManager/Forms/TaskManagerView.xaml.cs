using Autodesk.Revit.UI;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Common;
using KPLN_TaskManager.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;

namespace KPLN_TaskManager.Forms
{
    public partial class TaskManagerView : System.Windows.Controls.Page, IDockablePaneProvider, INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<TaskItemEntity> _collection;
        private DBProject _dBProject;

        private string _searchHeader;
        private string _selectedOpenStausTasks = "Open";
        private bool _fullTaskCollection = false;
        private string _subDepDependence = "AllSuDepTask";

        public TaskManagerView()
        {
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

            _dBProject = Module.CurrentDBProject;

            _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProject(Module.CurrentDBProject));

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
                
                bool isInputSubDepTask = MainDBService.CurrentDBUserSubDepartment.Id == task.DelegatedDepartmentId ? true : false;
                bool isOutputSubDepTask = MainDBService.CurrentDBUserSubDepartment.Id == task.CreatedTaskDepartmentId ? true : false;

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
            System.Windows.Controls.Button btn = sender as System.Windows.Controls.Button;
            if (btn.DataContext is TaskItemEntity taskItemEntity)
            {
                // Обновляю сущность с БД, чтобы получить АКТУАЛЬНУЮ инфу (если кто-то исправил, а я открыл после записи в БД)
                TaskItemEntity updatedTaskItemEntity = TaskManagerDBService.GetEntity_ByEntityId(taskItemEntity.Id);

                // Открываю окно
                TaskItemView taskItemView = new TaskItemView(updatedTaskItemEntity);
                taskItemView.ShowDialog();

                LoadTaskData();
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

            TaskItemEntity taskItemEntity = new TaskItemEntity(Module.CurrentDBProject.Id, MainDBService.CurrentDBUser.Id, MainDBService.CurrentDBUserSubDepartment.Id);
            TaskItemView taskItemView = new TaskItemView(taskItemEntity);
            bool? ceateViewResult = taskItemView.ShowDialog();
            if ((bool)ceateViewResult)
            {
                _collection.Add(taskItemView.CurrentTaskItemEntity);

                FilteredTasks = CollectionViewSource.GetDefaultView(_collection);
                FilteredTasks?.Refresh();
            }
        }

        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            string path = ExportToExcelService.SetPath();
            if (!string.IsNullOrEmpty(path))
            {
                List<TaskItemEntity> filteredTasks = new List<TaskItemEntity>();
                foreach (TaskItemEntity task in FilteredTasks)
                {
                    filteredTasks.Add(task);
                }

                ExportToExcelService.Run(path, _dBProject.Name, filteredTasks);
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
            if (MainDBService.CurrentDBUserSubDepartment.Id == 8)
            {
                switch (MainDBService.CurrentDBUser.Surname)
                {
                    case "Куцко":
                        return true;
                    case "Коломиец":
                        goto case "Федосеева";
                    case "Федосеева":
                        return task.DelegatedDepartmentId == 2 || task.CreatedTaskDepartmentId == 2;
                    case "Тарчоков":
                        goto case "Ямковой";
                    case "Ямковой":
                        return task.DelegatedDepartmentId == 3 || task.CreatedTaskDepartmentId == 3;
                    case "Садовская":
                        return task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                            || task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                            || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21
                            || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22;
                    case "Чичева":
                        return task.DelegatedDepartmentId == 4 || task.CreatedTaskDepartmentId == 4
                            || task.DelegatedDepartmentId == 6 || task.CreatedTaskDepartmentId == 6
                            || task.DelegatedDepartmentId == 20 || task.CreatedTaskDepartmentId == 20;
                }
            }
            //Для руководителей - верстка на лету под их отделы, чтобы они видели замечания по своим отделам
            if (MainDBService.CurrentDBUser.Surname == "Кудрова")
                return task.DelegatedDepartmentId == 4 || task.CreatedTaskDepartmentId == 4
                    || task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                    || task.DelegatedDepartmentId == 20 || task.CreatedTaskDepartmentId == 20
                    || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21;
            else if (MainDBService.CurrentDBUser.Surname == "Тамарин")
                return task.DelegatedDepartmentId == 5 || task.CreatedTaskDepartmentId == 5
                    || task.DelegatedDepartmentId == 21 || task.CreatedTaskDepartmentId == 21;
            else if (MainDBService.CurrentDBUser.Surname == "Колодий")
                return task.DelegatedDepartmentId == 6 || task.CreatedTaskDepartmentId == 6
                    || task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                    || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22;
            else if (MainDBService.CurrentDBUser.Surname == "Алиев")
                return task.DelegatedDepartmentId == 7 || task.CreatedTaskDepartmentId == 7
                    || task.DelegatedDepartmentId == 22 || task.CreatedTaskDepartmentId == 22;

            return true;
        }
    }
}
