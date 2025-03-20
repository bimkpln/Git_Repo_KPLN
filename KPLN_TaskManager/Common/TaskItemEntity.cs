using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Services;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace KPLN_TaskManager.Common
{
    public enum TaskStatusEnum
    {
        Open,
        Close
    }

    public class TaskItemEntity : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private int _projectId;
        private int _createdTaskUserId;
        private int _createdTaskDepartmentId;
        private string _header;
        private string _body;
        private int _delegatedDepartmentId = -1;
        public int _delegatedTaskUserId = -1;
        public int _bitrixParentTaskId = -1;
        private byte[] _imageBuffer;
        private int _bitrixTaskId;
        private string _modelName;
        private string _elementIds;
        private TaskStatusEnum _taskStatus;
        private string _createdTaskData;
        private string _lastChangeData;

        private SolidColorBrush _fill;
        private ImageSource _imageSource;

        /// <summary>
        /// Конструктор для Dapper (он по-умолчанию использует его, когда мапит данные из БД)
        /// </summary>
        public TaskItemEntity()
        {
        }

        public TaskItemEntity(int prjId, int createdUseId, int createdUserDepId) : this()
        {
            ProjectId = prjId;
            CreatedTaskUserId = createdUseId;
            CreatedTaskDepartmentId = createdUserDepId;
        }

        public TaskItemEntity(int prjId, int createdUseId, int createdUserDepId, string header, string body) : this(prjId, createdUseId, createdUserDepId)
        {
            TaskHeader = header;
            TaskBody = body;
        }

        #region Данные из БД
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Проект, к которому относится замечание
        /// </summary>
        public int ProjectId
        {
            get => _projectId;
            set
            {
                if (_projectId != value)
                {
                    _projectId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// ID пользователя, который создал/изменил замечание
        /// </summary>
        public int CreatedTaskUserId
        {
            get => _createdTaskUserId;
            set
            {
                if (_createdTaskUserId != value)
                {
                    _createdTaskUserId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Отдел, который создал/изменил замечание
        /// </summary>
        public int CreatedTaskDepartmentId
        {
            get => _createdTaskDepartmentId;
            set
            {
                if (_createdTaskDepartmentId != value)
                {
                    _createdTaskDepartmentId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Заголовок замечания
        /// </summary>
        public string TaskHeader
        {
            get => _header;
            set
            {
                _header = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TaskTitle));
            }
        }

        /// <summary>
        /// Отдел, которому адресовано замечание
        /// </summary>
        public int DelegatedDepartmentId
        {
            get => _delegatedDepartmentId;
            set
            {
                if (_delegatedDepartmentId != value)
                {
                    _delegatedDepartmentId = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(TaskTitle));
                    OnPropertyChanged(nameof(DBUserColl));
                }
            }
        }

        /// <summary>
        /// ID пользователя, который должен устранять замечание
        /// </summary>
        public int DelegatedTaskUserId
        {
            get => _delegatedTaskUserId;
            set
            {
                if (_delegatedTaskUserId != value)
                {
                    _delegatedTaskUserId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// ID родительской задачи в Bitrix
        /// </summary>
        public int BitrixParentTaskId
        {
            get => _bitrixParentTaskId;
            set
            {
                if (_bitrixParentTaskId != value)
                {
                    _bitrixParentTaskId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// ID задачи в Bitrix
        /// </summary>
        public int BitrixTaskId
        {
            get => _bitrixTaskId;
            set
            {
                if (value != _bitrixTaskId)
                {
                    _bitrixTaskId = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Текстовое сопровождение замечания
        /// </summary>
        public string TaskBody
        {
            get => _body;
            set
            {
                if (value != _body)
                {
                    _body = value;
                    OnPropertyChanged();
                }
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
                if (value != _imageBuffer)
                {
                    _imageBuffer = value;

                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ImageSource));
                }
            }
        }

        /// <summary>
        /// Имя файла, в котором находятся выбранные элементы
        /// </summary>
        public string ModelName
        {
            get => _modelName;
            set
            {
                _modelName = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Коллекция ID элементов ревит-модели, в формате, пригодном для обработки ревитом (11111,11112,11113....)
        /// </summary>
        public string ElementIds
        { 
            get => _elementIds;
            set
            {
                _elementIds = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ElementIdsCount));
            }
        }

        /// <summary>
        /// Статус элемента
        /// </summary>
        public TaskStatusEnum TaskStatus
        {
            get => _taskStatus;
            set
            {
                switch (value)
                {
                    case TaskStatusEnum.Close:
                        TaskBackground = new SolidColorBrush(Color.FromArgb(255, 0, 190, 104));
                        break;
                    case TaskStatusEnum.Open:
                        TaskBackground = new SolidColorBrush(Color.FromArgb(255, 255, 84, 42));
                        break;
                }

                _taskStatus = value;

                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Дата создания
        /// </summary>
        public string CreatedTaskData
        {
            get => _createdTaskData;
            set
            {
                if (value != _createdTaskData)
                {
                    _createdTaskData = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// Дата последнего редактирования
        /// </summary>
        public string LastChangeData
        {
            get => _lastChangeData;
            set
            {
                if (value != _lastChangeData)
                {
                    _lastChangeData = value;
                    OnPropertyChanged();
                }
            }
        }
        #endregion

        #region Дополнительная визуализация
        /// <summary>
        /// Заголовок замечания
        /// </summary>
        public string TaskTitle
        {
            get
            {
                if (CreatedTaskDepartmentId == 0 || DelegatedDepartmentId == 0)
                    return "Новое задание";
                else
                    return $"{CreatedTaskDepartmentCode} для {DelegatedDepartmentCode}: {TaskHeader}";
            }
        }

        /// <summary>
        /// Имя отдела, от которого задание исходит
        /// </summary>
        public string CreatedTaskDepartmentCode
        {
            get
            {
                string crDepCode = string.Empty;
                DBSubDepartment dBSubDepartment = MainDBService.DBSubDepartmentColl.FirstOrDefault(sd => sd.Id == CreatedTaskDepartmentId);
                if (dBSubDepartment != null)
                    crDepCode = dBSubDepartment.Code;

                return crDepCode;
            }
        }

        /// <summary>
        /// Имя отдела, которому делегировано замечание
        /// </summary>
        public string DelegatedDepartmentCode
        {
            get
            {
                string delegDepCode = string.Empty;
                DBSubDepartment dBSubDepartment = MainDBService.DBSubDepartmentColl.FirstOrDefault(sd => sd.Id == DelegatedDepartmentId);
                if (dBSubDepartment != null)
                    delegDepCode = dBSubDepartment.Code;
                
                return delegDepCode;
            }
        }

        public string CreatedTaskUserFullName
        {
            get => $"{MainDBService.UserDbService.GetDBUser_ById(CreatedTaskUserId).Surname} {MainDBService.UserDbService.GetDBUser_ById(CreatedTaskUserId).Name}";
        }

        /// <summary>
        /// Фон элемента
        /// </summary>
        public SolidColorBrush TaskBackground
        {
            get { return _fill; }
            set
            {
                _fill = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Изображение элемента
        /// </summary>
        public ImageSource ImageSource
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
                    _imageSource = bitmapImage;

                    return _imageSource;
                }
            }
            set
            {
                _imageSource = value;
                OnPropertyChanged();
            }
        }
        #endregion

        #region Дополнительные данные
        /// <summary>
        /// Коллекция отделов КПЛН
        /// </summary>
        public DBSubDepartment[] DBSubDepartmentColl
        {
            get => MainDBService.DBSubDepartmentColl;
        }

        /// <summary>
        /// Коллекция пользователей КПЛН
        /// </summary>
        public List<DBUser> DBUserColl
        {
            get => MainDBService.GetDBUsers_BySubDepId(DelegatedDepartmentId);
        }

        /// <summary>
        /// Счетчик количества элементов прикрепленных к задаче
        /// </summary>
        public int ElementIdsCount
        {
            get
            {
                if (ElementIds == null || !ElementIds.Any())
                    return 0;
                else
                    return Regex.Matches(ElementIds, ",").Count + 1;
            }
        }
        #endregion

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
