using KPLN_Library_Bitrix24Worker;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Data;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class UserSearch : Window
    {
        private readonly VM_UserEntity[] _collection;
        private readonly ListCollectionView _usersView;
        private bool _isOpeningProfile;

        public UserSearch(IEnumerable<VM_UserEntity> collection,
            Dictionary<string, List<string>> selectionRoles = null, string selectionInfo = null, int selectedElementCount = 0)
        {
            // Материализация исключает повторные запросы отделов в БД при каждом вводе.
            _collection = collection.ToArray();
            InitializeComponent();

            var rows = _collection.Select(user => new UserSearchRow(user, user.DBRevitUserName, selectionRoles)).ToList();
            if (selectionRoles != null)
            {
                var knownNames = new HashSet<string>(_collection.Select(user => user.DBRevitUserName?.Trim()),
                    StringComparer.OrdinalIgnoreCase);
                foreach (string name in selectionRoles.Keys.Where(name => !knownNames.Contains(name)))
                    rows.Add(new UserSearchRow(null, name, selectionRoles));

                Width = 780;
                MinWidth = 700;
                SelectionPanel.Visibility = Visibility.Visible;
                SelectionInfo.Text = selectionInfo;
                OnlySelection.Content = selectedElementCount > 1
                    ? "Только пользователи выбранных элементов"
                    : "Только пользователи выбранного элемента";
                OnlySelection.IsChecked = true;
            }

            _usersView = new ListCollectionView(rows);
            ApplyFilter();
            Users.ItemsSource = _usersView;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public IEnumerable<VM_UserEntity> Collection { get { return _collection; } }

        private void HandlePressBtn(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            Keyboard.Focus(UserInput);
        }

        private void UserInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void IncludeFired_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            // События XAML могут сработать до завершения InitializeComponent.
            if (_usersView == null)
                return;

            string searchName = UserInput.Text.Trim();
            bool includeFired = IncludeFired.IsChecked == true;
            bool onlySelection = OnlySelection.IsChecked == true;

            _usersView.Filter = item =>
            {
                UserSearchRow row = (UserSearchRow)item;
                if (!includeFired && row.User != null && row.User.DBUser.IsFired)
                    return false;

                if (onlySelection && string.IsNullOrEmpty(row.Roles))
                    return false;

                return searchName.Length == 0
                    || (!string.IsNullOrEmpty(row.DBRevitUserName)
                        && row.DBRevitUserName.IndexOf(searchName, StringComparison.OrdinalIgnoreCase) >= 0);
            };
            SearchStatus.Text = _usersView.Count == 0
                ? "Совпадений нет. Измените имя или настройки поиска."
                : string.Empty;
            SearchStatus.Visibility = _usersView.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void Users_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isOpeningProfile
                || !(e.OriginalSource is DependencyObject source)
                || !(ItemsControl.ContainerFromElement(Users, source) is ListBoxItem item)
                || !(item.DataContext is UserSearchRow row)
                || row.User == null)
                return;

            VM_UserEntity vmUserEntity = row.User;

            _isOpeningProfile = true;
            try
            {
                int id = await BitrixMessageSender.GetDBUserBitrixId_ByDBUser(vmUserEntity.DBUser);
                if (id <= 0)
                    throw new InvalidOperationException("Профиль пользователя в Bitrix24 не найден.");

                Process.Start(new ProcessStartInfo($@"https://kpln.bitrix24.ru/company/personal/user/{id}/")
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Не удалось открыть профиль Bitrix24.\n\n{ex.Message}",
                    "KPLN: Поиск пользователя", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            finally
            {
                _isOpeningProfile = false;
            }
        }
    }

    /// <summary>Строка поиска, включая логины выделения, отсутствующие в справочнике.</summary>
    public sealed class UserSearchRow
    {
        public UserSearchRow(VM_UserEntity user, string revitName, Dictionary<string, List<string>> roles)
        {
            User = user;
            DBRevitUserName = revitName;
            ShowRoles = roles != null;
            if (roles != null && revitName != null && roles.TryGetValue(revitName.Trim(), out List<string> labels))
                Roles = string.Join("\n", labels);
        }

        public VM_UserEntity User { get; }
        public string DBRevitUserName { get; }
        public string DBSystemName => User?.DBSystemName;
        public string UserDataView => User?.UserDataView ?? "Не найден в справочнике сотрудников";
        public string Roles { get; }
        public bool ShowRoles { get; }
    }
}

