using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class SendMsgToBitrix_VM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private readonly ICollectionView _userEntitysCollectionView;

        private string _filterBySurname;
        private string _filterByName;
        private string _filterByDepartment;

        private string _messageToSend_MainData;
        private string _messageToSend_UserComment;
        private byte[] _imageBuffer;
        private ImageSource _msgImageSource;

        public SendMsgToBitrix_VM(ObservableCollection<SendMsgToBitrix_UserEntity> userEntitysCollection)
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
        /// Изображение элемента
        /// </summary>
        public ImageSource MsgImageSource
        {
            get
            {
                if (ImageBuffer == null || ImageBuffer.Length == 0)
                    return null;

                using (MemoryStream ms = new MemoryStream(ImageBuffer))
                {
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = ms;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    _msgImageSource = bitmapImage;

                    return _msgImageSource;
                }
            }
            set
            {
                _msgImageSource = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Изображение элемента (массив байтов (BLOB))
        /// </summary>
        public byte[] ImageBuffer 
        {
            get => _imageBuffer;
            set
            {
                _imageBuffer = value;
                OnPropertyChanged(nameof(MsgImageSource));
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

        private void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
