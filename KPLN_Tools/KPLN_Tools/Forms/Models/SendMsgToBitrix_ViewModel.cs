using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;

namespace KPLN_Tools.Forms.Models
{
    public class SendMsgToBitrix_ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly ICollectionView _userEntitysCollectionView;
        
        private string _filterBySurname;
        private string _filterByName;
        private string _filterByDepartment;

        private string _messageToSend_MainData;
        private string _messageToSend_UserComment;

        public SendMsgToBitrix_ViewModel(ObservableCollection<SendMsgToBitrix_UserEntity> userEntitysCollection)
        {
            UserEntitysCollection = userEntitysCollection;

            // Создаем CollectionView для фильтрации
            _userEntitysCollectionView = CollectionViewSource.GetDefaultView(UserEntitysCollection);
            _userEntitysCollectionView.Filter = FilterUsers;
        }

        /// <summary>
        /// Коллекция элементов, которые будут отображаться в окне
        /// </summary>
        public ObservableCollection<SendMsgToBitrix_UserEntity> UserEntitysCollection { get; private set; }

        /// <summary>
        /// Выбранная коллекция элементов
        /// </summary>
        public IEnumerable<SendMsgToBitrix_UserEntity> SelectedElements
        {
            get => UserEntitysCollection.Where(ee => ee.IsSelected).ToList();
        }

        /// <summary>
        /// Свойство для фильтрации по фамилии
        /// </summary>
        public string FilterBySurname
        {
            get => _filterBySurname;
            set
            {
                _filterBySurname = value;
                OnPropertyChanged();

                _userEntitysCollectionView.Refresh(); 
            }
        }

        /// <summary>
        /// Свойство для фильтрации по имени
        /// </summary>
        public string FilterByName
        {
            get => _filterByName;
            set
            {
                _filterByName = value;
                OnPropertyChanged();

                _userEntitysCollectionView.Refresh(); 
            }
        }

        /// <summary>
        /// Свойство для фильтрации по отделу
        /// </summary>
        public string FilterByDepartment
        {
            get => _filterByDepartment;
            set
            {
                _filterByDepartment = value;
                OnPropertyChanged();

                _userEntitysCollectionView.Refresh();
            }
        }

        /// <summary>
        /// Основные данные по элементу для отправки сообщения
        /// </summary>
        public string MessageToSend_MainData
        {
            get => _messageToSend_MainData;
            set
            {
                if (value != _messageToSend_MainData)
                {
                    _messageToSend_MainData = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Пользовательский комментарий для отправки сообщения
        /// </summary>
        public string MessageToSend_UserComment
        {
            get => _messageToSend_UserComment;
            set
            {
                if (value != _messageToSend_UserComment)
                {
                    _messageToSend_UserComment = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Метод фильтрации
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool FilterUsers(object item)
        {
            if (item is SendMsgToBitrix_UserEntity user)
            {
                //Фильтруем по фамилии и имени
                bool surnameMatches = string.IsNullOrEmpty(FilterBySurname) || user.DBSurname.IndexOf(FilterBySurname, StringComparison.OrdinalIgnoreCase) >= 0;
                bool nameMatches = string.IsNullOrEmpty(FilterByName) || user.DBName.IndexOf(FilterByName, StringComparison.OrdinalIgnoreCase) >= 0;
                bool depMatches = string.IsNullOrEmpty(FilterByDepartment) || user.DBSubDepartment.Code.IndexOf(FilterByDepartment, StringComparison.OrdinalIgnoreCase) >= 0;

                return surnameMatches && nameMatches && depMatches;
            }

            return false;
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
