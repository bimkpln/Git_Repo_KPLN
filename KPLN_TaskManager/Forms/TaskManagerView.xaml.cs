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
            _dBProject = Module.CurrentDBProject;
            
            // Для бим-отдела верстка на лету под спецов, чтобы они видели замечания по своим отделам
            if (MainDBService.CurrentDBUserSubDepartment.Id == 8)
            {
                switch (MainDBService.CurrentDBUser.Surname)
                {
                    case "Куцко":
                        goto default;
                    case "Коломиец":
                        goto case "Федосеева";
                    case "Федосеева":
                        _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 2));
                        break;
                    case "Тарчоков":
                        goto case "Ямковой";
                    case "Ямковой":
                        _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 3));
                        break;
                    case "Садовская":
                        _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 5));
                        IEnumerable<TaskItemEntity> secondColl = TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 7);
                        foreach (TaskItemEntity item in secondColl)
                        {
                            _collection.Add(item);
                        }
                        break;
                    case "Чичева":
                        _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 4));
                        IEnumerable<TaskItemEntity> secondColl2 = TaskManagerDBService.GetEntities_ByDBProjectAndBIMUserSubDepId(Module.CurrentDBProject, 8, 6);
                        foreach (TaskItemEntity item in secondColl2)
                        {
                            _collection.Add(item);
                        }
                        break;
                    default:
                        _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProject(Module.CurrentDBProject));
                        break;
                }
            }
            // Для остальных пользователей - выбор из ЗВ и ЗИ
            else
                _collection = new ObservableCollection<TaskItemEntity>(TaskManagerDBService.GetEntities_ByDBProjectAndUser(Module.CurrentDBProject, MainDBService.CurrentDBUser));

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

                string inputSubDepTask = MainDBService.CurrentDBUserSubDepartment.Id == task.DelegatedDepartmentId ? "InputSubDepTask" : "Nope";
                string outputSubDepTask = MainDBService.CurrentDBUserSubDepartment.Id == task.CreatedTaskDepartmentId ? "OutputSubDepTask" : "Nope";
                bool matchesSubDep = SubDepDependence == "AllSuDepTask" || SubDepDependence.Equals(inputSubDepTask) || SubDepDependence.Equals(outputSubDepTask);

                return matchesTitle && matchesStatus && matchesSubDep;
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
    }
}
