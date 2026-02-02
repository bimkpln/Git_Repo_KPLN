using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace KPLN_DefaultPanelExtension_Modify.Forms.Models
{
    public sealed class SendMsgToBitrix_UserEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _userDataView;

        private bool _isSelected = false;

        public SendMsgToBitrix_UserEntity(DBUser dbUser, DBSubDepartment dbSubDepartment)
        {
            DBUser = dbUser;
            DBSubDepartment = dbSubDepartment;

            DBName = dbUser.Name;
            DBSurname = dbUser.Surname;
            DBSystemName = dbUser.SystemName;
            DBRevitUserName = dbUser.RevitUserName;

            UserDataView = $"{DBSurname} {DBName}. Отдел: {DBSubDepartment.Code}";
        }

        /// <summary>
        /// Пользователь из БД
        /// </summary>
        public DBUser DBUser { get; private set; }

        /// <summary>
        /// Пользователь из БД
        /// </summary>
        public DBSubDepartment DBSubDepartment { get; private set; }

        /// <summary>
        /// Имя пользователя
        /// </summary>
        public string DBName { get; private set; }

        /// <summary>
        /// Фамилия пользователя
        /// </summary>
        public string DBSurname { get; private set; }

        /// <summary>
        /// Имя учетки Windows
        /// </summary>
        public string DBSystemName { get; private set; }

        /// <summary>
        /// Имя учетки Revit
        /// </summary>
        public string DBRevitUserName { get; private set; }

        /// <summary>
        /// Данные, выводимые к окно
        /// </summary>
        public string UserDataView
        {
            get => _userDataView;
            private set
            {
                if (value != _userDataView)
                {
                    _userDataView = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Для окна wpf - пометка, что элемент выбран
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
