using Autodesk.Revit.DB;
using KPLN_Library_SQLiteWorker;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using KPLN_TaskManager.Services;
using System;
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

    public sealed class TaskItemEntity : INotifyPropertyChanged
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
        public string _pathToImageBufferDB;
        private int _bitrixTaskId;
        private string _modelName;
        private long _modelViewId;
        private string _elementIds;
        private TaskStatusEnum _taskStatus;
        private string _createdTaskData;
        private string _lastChangeData;
        private ImageSource _imageSource;

        
        private int _te_ImageBuffer_Current = 0;
        private List<TaskEntity_ImageBuffer> _teImageBuffer = new List<TaskEntity_ImageBuffer>();
        private SolidColorBrush _fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 190, 104));

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
        /// Отдел, ОТ имени которого создал/изменил замечание
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
                    OnPropertyChanged(nameof(TaskTitle));
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
                    OnPropertyChanged(nameof(ModelNamesColl));

                    if (Id == 0)
                        FindModelByDlgSubDep();
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
        /// Ссылка на БД отчетов
        /// </summary>
        public string PathToImageBufferDB
        {
            get => _pathToImageBufferDB;
            set
            {
                _pathToImageBufferDB = value;
                OnPropertyChanged();
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
        /// ID вида, на котором были выбраны элементы
        /// </summary>
        public long ModelViewId
        {
            get => _modelViewId;
            set
            {
                _modelViewId = value;
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
                        TaskBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 0, 190, 104));
                        break;
                    case TaskStatusEnum.Open:
                        TaskBackground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 255, 84, 42));
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

        /// <summary>
        /// Изображение элемента
        /// </summary>
        public ImageSource TaskImageSource
        {
            get
            {
                if (TE_ImageBufferColl.All(buff => buff.ImageBuffer == null || buff.ImageBuffer.Length == 0))
                    return null;

                using (MemoryStream ms = new MemoryStream(TE_ImageBufferColl[TE_ImageBuffer_Current].ImageBuffer))
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

        #region Дополнительная визуализация
        /// <summary>
        /// Индекс текущего TE_ImageBuffer
        /// </summary>
        public int TE_ImageBuffer_Current
        {
            get => _te_ImageBuffer_Current;
            set
            {
                _te_ImageBuffer_Current = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TaskImageSource));
                OnPropertyChanged(nameof(CurrentImgSpecialFormat));
            }
        }

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
                DBSubDepartment dBSubDepartment = DBMainService.DBSubDepartmentColl.FirstOrDefault(sd => sd.Id == CreatedTaskDepartmentId);
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
                DBSubDepartment dBSubDepartment = DBMainService.DBSubDepartmentColl.FirstOrDefault(sd => sd.Id == DelegatedDepartmentId);
                if (dBSubDepartment != null)
                    delegDepCode = dBSubDepartment.Code;

                return delegDepCode;
            }
        }

        public string CreatedTaskUserFullName
        {
            get
            {
                DBUser user = DBMainService.UserDbService.GetDBUser_ById(CreatedTaskUserId);
                return $"{user.Surname} {user.Name}";
            }
        }

        /// <summary>
        /// Фон элемента
        /// </summary>
        public SolidColorBrush TaskBackground
        {
            get => _fill;
            set
            {
                _fill = value;
                OnPropertyChanged();
            }
        }

        public string CurrentImgSpecialFormat
        {
            get => $"{TE_ImageBuffer_Current + 1}/{TE_ImageBufferColl.Count}";
        }
        #endregion

        #region Дополнительные данные
        /// <summary>
        /// Контейнер изображений для замечания (спец-класс)
        /// </summary>
        public List<TaskEntity_ImageBuffer> TE_ImageBufferColl
        {
            get
            {
                if (_teImageBuffer.Count() > 0)
                    return _teImageBuffer;

                if (this.Id != 0 && !string.IsNullOrWhiteSpace(this.PathToImageBufferDB))
                    _teImageBuffer = TM_IBDBService.GetEntity_ByEntityId(this)
                        ?? new List<TaskEntity_ImageBuffer>() { new TaskEntity_ImageBuffer() };

                return _teImageBuffer;
            }
            set
            {
                _teImageBuffer = value;

                // Обновляю изображение в окне
                using (MemoryStream ms = new MemoryStream(_teImageBuffer[TE_ImageBuffer_Current].ImageBuffer))
                {
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = ms;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    TaskImageSource = bitmapImage;
                }

                OnPropertyChanged(nameof(CurrentImgSpecialFormat));
            }
        }

        /// <summary>
        /// Коллекция отделов КПЛН  ОТ
        /// </summary>
        public DBSubDepartment[] DBSubDepartmentColl_Output
        {
            get
            {
                return DBMainService
                 .DBSubDepartmentColl
                 .Where(sd =>
                     sd.Id == DBMainService.CurrentUserDBSubDepartment.Id
                     || sd.DependentSubDepId == DBMainService.CurrentUserDBSubDepartment.Id)
                 .ToArray();
            }
        }

        /// <summary>
        /// Коллекция отделов КПЛН  ДЛЯ 
        /// </summary>
        public DBSubDepartment[] DBSubDepartmentColl_Input
        {
            get => DBMainService.DBSubDepartmentColl;
        }

        /// <summary>
        /// Коллекция пользователей КПЛН
        /// </summary>
        public List<DBUser> DBUserColl
        {
            get
            {
                if (DelegatedDepartmentId == 0 || DelegatedDepartmentId == -1)
                    return new List<DBUser>(1) { new DBUser() { Id = -2, Surname = "<Сначала выбери отдел>" } };

                List<DBUser> delSubDepUserColl = DBMainService
                    .UserDbService
                    .GetDBUsers_BySubDepID(DelegatedDepartmentId)
                    .OrderBy(x => x.Surname)
                    .ToList();

                return delSubDepUserColl;
            }
        }

        /// <summary>
        /// Коллекция моделей в открытом документе
        /// </summary>
        public HashSet<string> ModelNamesColl
        {
            get
            {
                HashSet<string> result = new HashSet<string>() { "<Весь проект>" };

                // Если НЕ новое замечание - добавляю имя
                if (Id > 0 && !string.IsNullOrEmpty(ModelName))
                    result.Add(ModelName);

                // Если отдел делегирования совпадает с отделом открытой модели - добавляю в списк эту же модель
                if (Module.CurrentDoc != null && Module.CurrnetDocSubDep.Id == DelegatedDepartmentId)
                    result.Add(CurrentModelName(Module.CurrentDoc));

                // Заполняю список линками в зависимости от выбранного отдела делегирования
                if (Module.CurrentDoc != null && DelegatedDepartmentId > 0)
                {
                    foreach (string docName in GetDocRLinkNamesByDelegatedDep(DelegatedDepartmentId))
                    {
                        result.Add(docName);
                    }
                }

                return result;
            }
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

        /// <summary>
        /// Выделить имя файла из Document
        /// </summary>
        /// <returns></returns>
        internal static string CurrentModelName(Document doc)
        {
            string openViewFileName = doc.IsWorkshared && !doc.IsDetached
                ? ModelPathUtils.ConvertModelPathToUserVisiblePath(doc.GetWorksharingCentralModelPath())
                : doc.PathName;

            if (openViewFileName.ToLower().Contains("rsn"))
                return openViewFileName.Split('/').FirstOrDefault(fn => fn.Contains(".rvt"));
            else
                return openViewFileName.Split('\\').FirstOrDefault(fn => fn.Contains(".rvt"));
        }

        /// <summary>
        /// Подобрать имя модели, в зависимости от отдела, которому задача делегируется
        /// </summary>
        internal void FindModelByDlgSubDep()
        {
            // Если отдел ответсвенного совпадает с отделом модели - то ставим на этот же файл
            if (Module.CurrnetDocSubDep.Id == DelegatedDepartmentId)
                ModelName = ModelNamesColl.FirstOrDefault(name => name.Contains(CurrentModelName(Module.CurrentDoc)));
            else
                ModelName = ModelNamesColl.FirstOrDefault();
        }

        private static string[] GetDocRLinkNamesByDelegatedDep(int depId)
        {
            if (Module.CurrentDoc == null)
                return null;

            RevitLinkInstance[] rlInsts = new FilteredElementCollector(Module.CurrentDoc)
                .OfClass(typeof(RevitLinkInstance))
                .WhereElementIsNotElementType()
                .Cast<RevitLinkInstance>()
                .ToArray();

            HashSet<string> rlNames = new HashSet<string>(rlInsts.Length);
            foreach (RevitLinkInstance inst in rlInsts)
            {
                string instName = inst.Name;
                DBSubDepartment rlSubDep = DBMainService.SubDepartmentDbService.GetDBSubDepartment_ByRevitDocFullPath(instName);
                if (rlSubDep.Id == depId)
                {
                    // Чистка от доп. параметров в имени линка
                    string instAttrSepar = " : ";
                    if (instName.Contains(instAttrSepar))
                        rlNames.Add(instName.Split(new string[] { instAttrSepar }, StringSplitOptions.None)[0]);
                    else
                        rlNames.Add(instName);
                }
            }

            return rlNames.ToArray();
        }

        private void OnPropertyChanged([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
