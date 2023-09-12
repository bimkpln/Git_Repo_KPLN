using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Net.Http;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace KPLN_Tools.Forms
{
    public partial class UserSearch : Window
    {
        private readonly IEnumerable<DBUser> _collection;

        public UserSearch(IEnumerable<DBUser> collection)
        {
            _collection = collection;
            InitializeComponent();

            Users.ItemsSource = _collection;
            PreviewKeyDown += new KeyEventHandler(HandlePressBtn);
        }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public IEnumerable<DBUser> Collection { get { return _collection; } }

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

            ObservableCollection<DBUser> filteredElement = new ObservableCollection<DBUser>();

            foreach (DBUser user in Collection)
            {
                if (user.SystemName.ToLower().Contains(_searchName))
                {
                    filteredElement.Add(user);
                }
            }

            Users.ItemsSource = filteredElement;
        }

        private async void Users_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            ListBox listBox = (ListBox)sender;
            if (!(listBox.SelectedItem is DBUser dBUser))
                throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД\n\n");

            string id = string.Empty;
            using (HttpClient client = new HttpClient())
            {
                // Выполнение GET - запроса к странице
                HttpResponseMessage response = await client.GetAsync(String.Format(@"https://kpln.bitrix24.ru/rest/152/rud1zqq5p9ol00uk/user.search.json?LAST_NAME={0}", $"{dBUser.Surname}"));
                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    if (string.IsNullOrEmpty(content))
                        throw new Exception("\n[KPLN]: Ошибка получения ответа от Bitrix\n\n");

                    dynamic dynDeserilazeData = JsonConvert.DeserializeObject<dynamic>(content);
                    dynamic responseResult = dynDeserilazeData.result;
                    id = responseResult[0].ID.ToString();
                }
            }

            if (id == string.Empty)
                throw new Exception("\n[KPLN]: Ошибка получения пользователя из БД - не удалось получить id-пользователя Bitrix\n\n");

            Process.Start(new ProcessStartInfo($@"https://kpln.bitrix24.ru/company/personal/user/{id}/")
            {
                UseShellExecute = true
            });
        }
    }
}
