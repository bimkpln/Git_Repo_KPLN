using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Runtime.CompilerServices;

namespace KPLN_TaskManager.Common
{
    public class TaskItemEntity_Comment : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _message;
        private string _createdMsgData;

        /// <summary>
        /// Конструктор для Dapper (он по-умолчанию использует его, когда мапит данные из БД)
        /// </summary>
        public TaskItemEntity_Comment()
        {
        }

        public TaskItemEntity_Comment(int taskId, int userId, string msg, string createdMsgData)
        {
            TaskItemEntityId = taskId;
            DBUserId = userId;
            Message = msg;
            CreatedMsgData = createdMsgData;
        }

        #region Данные из БД
        [Key]
        public int Id { get; set; }

        [ForeignKey(nameof(TaskItemEntity))]
        public int TaskItemEntityId { get; set; }

        public int DBUserId { get; set; }

        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CommentFullData));
            }
        }

        /// <summary>
        /// Дата создания
        /// </summary>
        public string CreatedMsgData
        {
            get => _createdMsgData;
            set
            {
                _createdMsgData = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CommentFullData));
            }
        }
        #endregion

        #region Дополнительная визуализация
        /// <summary>
        /// Вывод данных для пользователя
        /// </summary>
        public string CommentFullData
        {
            get => $"<{UserFullData} {CreatedMsgData}>: {Message}";
        }

        /// <summary>
        /// Данные по пользователю
        /// </summary>
        public string UserFullData
        {
            get => $"{CurrentDBUser.Surname} {CurrentDBUser.Name}";
        }
        #endregion

        #region Дополнительные данные
        /// <summary>
        /// Ссылка на выбранного пользователя из БД
        /// </summary>
        private DBUser CurrentDBUser
        {
            get => DBMainService.UserDbService.GetDBUser_ById(DBUserId);
        }
        #endregion

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
