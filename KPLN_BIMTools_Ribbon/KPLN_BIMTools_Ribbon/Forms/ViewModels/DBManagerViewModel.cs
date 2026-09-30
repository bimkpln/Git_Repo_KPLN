using KPLN_BIMTools_Ribbon.Forms.Commands;
using KPLN_BIMTools_Ribbon.Forms.Models;
using KPLN_Library_DBWorker;
using KPLN_Library_DBWorker.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace KPLN_BIMTools_Ribbon.Forms.ViewModels
{
    public sealed class DBManagerViewModel : INotifyPropertyChanged
    {
        private enum ManagerMode
        {
            ScenarioSelection,
            Projects,
            Users
        }

        private readonly DBManager _mainWindow;
        private ManagerMode _currentMode = ManagerMode.ScenarioSelection;
        private DBProject _selectedProject;
        private DBUserManagerItem _selectedUser;
        private string _projectFilterText;
        private string _userFilterText;
        private bool _showFiredUsers;

        public DBManagerViewModel(DBManager mainWindow)
        {
            _mainWindow = mainWindow;

            Projects = new ObservableCollection<DBProject>();
            Users = new ObservableCollection<DBUserManagerItem>();

            ProjectsView = CollectionViewSource.GetDefaultView(Projects);
            ProjectsView.Filter = FilterProject;
            UsersView = CollectionViewSource.GetDefaultView(Users);
            UsersView.Filter = FilterUser;

            OpenProjectsCommand = new RelayCommand(OpenProjects);
            OpenUsersCommand = new RelayCommand(OpenUsers);
            BackToScenariosCommand = new RelayCommand(BackToScenarios);
            RefreshProjectsCommand = new RelayCommand(() => LoadProjects(SelectedProject?.Id ?? -1));
            RefreshUsersCommand = new RelayCommand(() => LoadUsers(SelectedUser?.Id ?? -1));
            CreateProjectCommand = new RelayCommand(CreateProject);
            ToggleProjectAccessCommand = new RelayCommand(ToggleProjectAccess, () => SelectedProject != null);
            ToggleUserAccessCommand = new RelayCommand(ToggleUserAccess, () => SelectedUser != null);
            CloseWindowCommand = new RelayCommand(() => _mainWindow.Close());
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<DBProject> Projects { get; }

        public ObservableCollection<DBUserManagerItem> Users { get; }

        public ICollectionView ProjectsView { get; }

        public ICollectionView UsersView { get; }

        public ICommand OpenProjectsCommand { get; }

        public ICommand OpenUsersCommand { get; }

        public ICommand BackToScenariosCommand { get; }

        public ICommand RefreshProjectsCommand { get; }

        public ICommand RefreshUsersCommand { get; }

        public ICommand CreateProjectCommand { get; }

        public ICommand ToggleProjectAccessCommand { get; }

        public ICommand ToggleUserAccessCommand { get; }

        public ICommand CloseWindowCommand { get; }

        public bool IsScenarioSelectionVisible => _currentMode == ManagerMode.ScenarioSelection;

        public bool IsProjectsVisible => _currentMode == ManagerMode.Projects;

        public bool IsUsersVisible => _currentMode == ManagerMode.Users;

        public DBProject SelectedProject
        {
            get => _selectedProject;
            set
            {
                if (_selectedProject == value)
                    return;

                _selectedProject = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SelectedProjectActionText));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public DBUserManagerItem SelectedUser
        {
            get => _selectedUser;
            set
            {
                if (_selectedUser == value)
                    return;

                _selectedUser = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SelectedUserActionText));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string ProjectFilterText
        {
            get => _projectFilterText;
            set
            {
                if (_projectFilterText == value)
                    return;

                _projectFilterText = value;
                OnPropertyChanged();
                ProjectsView.Refresh();
            }
        }

        public string UserFilterText
        {
            get => _userFilterText;
            set
            {
                if (_userFilterText == value)
                    return;

                _userFilterText = value;
                OnPropertyChanged();
                UsersView.Refresh();
            }
        }

        public bool ShowFiredUsers
        {
            get => _showFiredUsers;
            set
            {
                if (_showFiredUsers == value)
                    return;

                _showFiredUsers = value;
                OnPropertyChanged();
                UsersView.Refresh();

                if (SelectedUser != null && !UsersView.Contains(SelectedUser))
                    SelectedUser = UsersView.Cast<DBUserManagerItem>().FirstOrDefault();
            }
        }

        public string SelectedProjectActionText =>
            SelectedProject != null && SelectedProject.IsClosed
                ? "Открыть проект"
                : "Закрыть проект";

        public string SelectedUserActionText =>
            SelectedUser != null && SelectedUser.IsUserRestricted
                ? "Открыть доступ"
                : "Закрыть доступ";

        private void OpenProjects()
        {
            SetMode(ManagerMode.Projects);
            LoadProjects();
        }

        private void OpenUsers()
        {
            SetMode(ManagerMode.Users);
            LoadUsers();
        }

        private void BackToScenarios()
        {
            SelectedProject = null;
            SelectedUser = null;
            SetMode(ManagerMode.ScenarioSelection);
        }

        private void SetMode(ManagerMode mode)
        {
            _currentMode = mode;
            OnPropertyChanged(nameof(IsScenarioSelectionVisible));
            OnPropertyChanged(nameof(IsProjectsVisible));
            OnPropertyChanged(nameof(IsUsersVisible));
        }

        private void LoadProjects(int selectedProjectId = -1)
        {
            try
            {
                DBProject[] projects = SQLiteMainService
                    .SQLitePrjServiceInst
                    .GetDBProjects_All()
                    .OrderBy(project => project.Id)
                    .ToArray();

                Projects.Clear();
                foreach (DBProject project in projects)
                    Projects.Add(project);

                ProjectsView.Refresh();
                SelectedProject = Projects.FirstOrDefault(project => project.Id == selectedProjectId)
                    ?? Projects.FirstOrDefault();
            }
            catch (Exception ex)
            {
                ShowError("Не удалось загрузить проекты из KPLN_Loader_MainDB.", ex);
            }
        }

        private void LoadUsers(int selectedUserId = -1)
        {
            try
            {
                Dictionary<int, string> subDepartments = SQLiteMainService
                    .SQLiteSubDepServiceInst
                    .GetDBSubDepartments()
                    .GroupBy(subDepartment => subDepartment.Id)
                    .ToDictionary(group => group.Key, group => group.First().Name);

                DBUserManagerItem[] users = SQLiteMainService
                    .SQLiteUserServiceInst
                    .GetDBUsers()
                    .Select(user => new DBUserManagerItem(
                        user,
                        subDepartments.TryGetValue(user.SubDepartmentId, out string name) ? name : null))
                    .OrderBy(user => user.Id)
                    .ToArray();

                Users.Clear();
                foreach (DBUserManagerItem user in users)
                    Users.Add(user);

                UsersView.Refresh();
                SelectedUser = UsersView
                    .Cast<DBUserManagerItem>()
                    .FirstOrDefault(user => user.Id == selectedUserId)
                    ?? UsersView.Cast<DBUserManagerItem>().FirstOrDefault();
            }
            catch (Exception ex)
            {
                ShowError("Не удалось загрузить пользователей из KPLN_Loader_MainDB.", ex);
            }
        }

        private void CreateProject()
        {
            DBProjectCreateWindow createWindow = new DBProjectCreateWindow
            {
                Owner = _mainWindow
            };

            if (createWindow.ShowDialog() == true)
                LoadProjects();
        }

        private void ToggleProjectAccess()
        {
            if (SelectedProject == null)
                return;

            DBProject project = SelectedProject;
            bool closeProject = !project.IsClosed;
            string newStatus = closeProject ? "закрыт" : "открыт";
            string prompt =
                $"Проект «{project.Code} {project.Stage} — {project.Name}» будет {newStatus}.\n\n" +
                (closeProject
                    ? "Пользователи без специального допуска не смогут работать с проектом."
                    : "Пользователи снова смогут работать с проектом.");

            if (MessageBox.Show(
                    _mainWindow,
                    prompt,
                    "KPLN_DB: изменение доступа к проекту",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                SQLiteMainService.SQLitePrjServiceInst.UpdateDBProject_IsClosed(project, closeProject);
                LoadProjects(project.Id);
            }
            catch (Exception ex)
            {
                ShowError("Не удалось изменить статус проекта.", ex);
            }
        }

        private void ToggleUserAccess()
        {
            if (SelectedUser == null)
                return;

            DBUserManagerItem userItem = SelectedUser;
            bool restrictUser = !userItem.IsUserRestricted;
            string newStatus = restrictUser ? "закрыт" : "открыт";
            string prompt =
                $"Доступ пользователя «{userItem.FullName}» ({userItem.SystemName}) будет {newStatus}.\n\n" +
                (restrictUser
                    ? "Пользователь не сможет работать с реальными проектами."
                    : "Ограничение на работу с реальными проектами будет снято.");

            if (userItem.Id == SQLiteMainService.CurrentDBUser?.Id && restrictUser)
                prompt += "\n\nВнимание: это текущий пользователь Revit.";

            if (MessageBox.Show(
                    _mainWindow,
                    prompt,
                    "KPLN_DB: изменение доступа пользователя",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning) != MessageBoxResult.Yes)
                return;

            try
            {
                SQLiteMainService.SQLiteUserServiceInst.UpdateDBUser_IsUserRestricted(
                    userItem.DBUser,
                    restrictUser);

                if (SQLiteMainService.CurrentDBUser?.Id == userItem.Id)
                    SQLiteMainService.CurrentDBUser.IsUserRestricted = restrictUser;

                LoadUsers(userItem.Id);
            }
            catch (Exception ex)
            {
                ShowError("Не удалось изменить доступ пользователя.", ex);
            }
        }

        private bool FilterProject(object item)
        {
            DBProject project = item as DBProject;
            if (project == null || string.IsNullOrWhiteSpace(ProjectFilterText))
                return project != null;

            string filter = ProjectFilterText.Trim();
            return Contains(project.Id.ToString(), filter)
                || Contains(project.Name, filter)
                || Contains(project.Code, filter)
                || Contains(project.Stage, filter)
                || Contains(project.RevitVersion.ToString(), filter)
                || Contains(project.MainPath, filter)
                || Contains(project.RevitServerPath, filter);
        }

        private bool FilterUser(object item)
        {
            DBUserManagerItem user = item as DBUserManagerItem;
            if (user == null)
                return false;

            if (!ShowFiredUsers && user.IsFired)
                return false;

            if (string.IsNullOrWhiteSpace(UserFilterText))
                return true;

            string filter = UserFilterText.Trim();
            return Contains(user.Id.ToString(), filter)
                || Contains(user.FullName, filter)
                || Contains(user.SystemName, filter)
                || Contains(user.SubDepartmentName, filter)
                || Contains(user.RevitUserName, filter)
                || Contains(user.BitrixUserID.ToString(), filter);
        }

        private static bool Contains(string value, string filter) =>
            !string.IsNullOrWhiteSpace(value)
            && value.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;

        private void ShowError(string text, Exception ex)
        {
            MessageBox.Show(
                _mainWindow,
                $"{text}\n\n{ex.Message}",
                "KPLN_DB: ошибка",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
