using Autodesk.Revit.DB;
using KPLN_Library_ExtensibleStorage;
using System;
using System.Windows.Media;

namespace KPLN_ModelChecker_Lib.Common
{
    /// <summary>
    /// Расширенная сущность ExtensibleStorage
    /// </summary>
    public sealed class ExtensibleStorageEntity
    {
        // Последний запуск
        private ExtensibleStorageBuilder _esBuilderRun;
        public readonly string LastRunFieldName = "Last_Run";

        // Комментарий по внесению в допустимое
        private ExtensibleStorageBuilder _esBuilderUserText;
        public readonly string UserTextFieldName = "Approve_Comment";

        // Ключевой комментарий - маркер
        private ExtensibleStorageBuilder _esBuildergMarker;
        public readonly string MarkerFieldName = "Main_Marker";

        /// <summary>
        /// Цвет текста при выводе в форме (если нужно переопределить)
        /// </summary>
        public SolidColorBrush TextColor { get; set; }

        /// <summary>
        /// Имя проверки
        /// </summary>
        public string CheckName { get; private set; }

        /// <summary>
        /// Имя основного Storage
        /// </summary>
        public string MainStorageName { get; private set; }

        /// <summary>
        /// Данные из Storage последнего запуска
        /// </summary>
        public string LastRunText { get; set; }

        /// <summary>
        /// GUID для Storage последнего запуска
        /// </summary>
        public Guid LastRunGuid { get; private set; }

        /// <summary>
        /// GUID для Storage комментария пользователя (для допустимых)
        /// </summary>
        public Guid UserTextGuid { get; private set; }

        /// <summary>
        /// GUID для Storage ключевого комментария
        /// </summary>
        public Guid MarkerGuid { get; private set; }

        /// <summary>
        /// Extensible Storage для последнего запуска
        /// </summary>
        public ExtensibleStorageBuilder ESBuilderRun
        {
            get
            {
                if (_esBuilderRun == null) _esBuilderRun = new ExtensibleStorageBuilder(LastRunGuid, LastRunFieldName, MainStorageName);
                return _esBuilderRun;
            }
        }

        /// <summary>
        /// Extensible Storage для пользовательского комментария
        /// </summary>
        public ExtensibleStorageBuilder ESBuilderUserText
        {
            get
            {
                if (_esBuilderUserText == null) _esBuilderUserText = new ExtensibleStorageBuilder(UserTextGuid, UserTextFieldName, MainStorageName);
                return _esBuilderUserText;
            }
        }

        /// <summary>
        /// Extensible Storage для ключевого комментария (помечает факт запуска проектировщиком)
        /// </summary>
        public ExtensibleStorageBuilder ESBuildergMarker
        {
            get
            {
                if (_esBuildergMarker == null) _esBuildergMarker = new ExtensibleStorageBuilder(MarkerGuid, MarkerFieldName, MainStorageName);
                return _esBuildergMarker;
            }
            set
            {
                _esBuildergMarker = value;
            }
        }

        /// <summary>
        /// Метод инициализации статических свойств классов для работы с ExtensibleStorage
        /// </summary>
        public ExtensibleStorageEntity(string checkName, string mainStorageName, Guid lastRunGuid)
        {
            CheckName = checkName;
            MainStorageName = mainStorageName;
            LastRunGuid = lastRunGuid;
        }

        /// <summary>
        /// Метод инициализации статических свойств классов для работы с ExtensibleStorage. Включая текстовую пометку от пользователя
        /// </summary>
        public ExtensibleStorageEntity(string checkName, string mainStorageName, Guid lastRunGuid, Guid userTextGuid) : this(checkName, mainStorageName, lastRunGuid)
        {
            UserTextGuid = userTextGuid;
        }

        /// <summary>
        /// Метод инициализации статических свойств классов для работы с ExtensibleStorage. Включая текстовую пометку от пользователя и основной маркер
        /// </summary>
        public ExtensibleStorageEntity(string checkName, string mainStorageName, Guid lastRunGuid, Guid userTextGuid, Guid markerGuid) : this(checkName, mainStorageName, lastRunGuid, userTextGuid)
        {
            MarkerGuid = markerGuid;
        }

        public static ErrorStatus SetApproveStatusByUserComment(object obj, ExtensibleStorageBuilder objESB, ErrorStatus ifNullComment)
        {
            ErrorStatus currentStatus;
            if (obj is Element elem)
            {
                if (objESB.IsDataExists_Text(elem))
                    currentStatus = ErrorStatus.Approve;
                else 
                    currentStatus = ifNullComment;
            }
            else
                throw new Exception($"{obj} - не Element Revit");


            return currentStatus;
        }
    }
}
