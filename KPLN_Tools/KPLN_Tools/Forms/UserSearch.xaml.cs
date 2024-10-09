using KPLN_Library_Bitrix24Worker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_Tools.Forms.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class UserSearch : Window
    {
        private readonly IEnumerable<VM_UserEntity> _collection;

        public UserSearch(IEnumerable<VM_UserEntity> collection)
        {
            _collection = collection;
            InitializeComponent();

            Users.ItemsSource = _collection;
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

        private void ScrollViewer_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Выполните здесь нужное вам поведение при прокрутке колесом мыши
            ScrollViewer.ScrollToVerticalOffset(ScrollViewer.VerticalOffset - e.Delta);
            // Пометьте событие как обработанное, чтобы оно не передалось другим элементам
            e.Handled = true;
        }

        private void UserInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;
            string _searchName = textBox.Text.ToLower();

            ObservableCollection<VM_UserEntity> filteredElement = new ObservableCollection<VM_UserEntity>();

            foreach (VM_UserEntity user in Collection)
            {
                string revitUserName = user.DBUser.RevitUserName;
                if (!string.IsNullOrEmpty(revitUserName) && revitUserName.ToLower().Contains(_searchName))
                    filteredElement.Add(user);
            }

            Users.ItemsSource = filteredElement;
        }

        private async void Users_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ListBox listBox = (ListBox)sender;
            if (!(listBox.SelectedItem is VM_UserEntity vmUserEntity))
                throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД\n\n");

            int id = await BitrixMessageSender.GetDBUserBitrixId_ByDBUser(vmUserEntity.DBUser);
            Process.Start(new ProcessStartInfo($@"https://kpln.bitrix24.ru/company/personal/user/{id}/")
            {
                UseShellExecute = true
            });
        }
    }
}
