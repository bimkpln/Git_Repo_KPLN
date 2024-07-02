using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.Forms;
using System;
using System.Windows.Media;

namespace KPLN_ModelChecker_User.Common
{
    public sealed class ExtensibleStorageEntity
    {
        // Последний запуск
        private ExtensibleStorageBuilder _esBuilderRun;
        internal readonly string LastRunFieldName = "Last_Run";

        // Комментарий по внесению в допустимое
        private ExtensibleStorageBuilder _esBuilderUserText;
        internal readonly string UserTextFieldName = "Approve_Comment";

        // Ключевой комментарий - маркер
        private ExtensibleStorageBuilder _esBuildergMarker;
        internal readonly string MarkerFieldName = "Main_Marker";

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
        internal string MainStorageName { get; private set; }

        /// <summary>
        /// Данные из Storage последнего запуска
        /// </summary>
        public string LastRunText { get; set; }

        /// <summary>
        /// GUID для Storage последнего запуска
        /// </summary>
        internal Guid LastRunGuid { get; private set; }

        /// <summary>
        /// GUID для Storage комментария пользователя (для допустимых)
        /// </summary>
        internal Guid UserTextGuid { get; private set; }

        /// <summary>
        /// GUID для Storage ключевого комментария
        /// </summary>
        internal Guid MarkerGuid { get; private set; }

        /// <summary>
        /// Extensible Storage для последнего запуска
        /// </summary>
        internal ExtensibleStorageBuilder ESBuilderRun
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
        internal ExtensibleStorageBuilder ESBuilderUserText
        {
            get
            {
                if (_esBuilderUserText == null) _esBuilderUserText = new ExtensibleStorageBuilder(UserTextGuid, UserTextFieldName, MainStorageName);
                return _esBuilderUserText;
            }
        }

        /// <summary>
        /// Extensible Storage для ключевого комментария
        /// </summary>
        internal ExtensibleStorageBuilder ESBuildergMarker
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
        internal ExtensibleStorageEntity(string checkName, string mainStorageName, Guid lastRunGuid)
        {
            CheckName = checkName;
            MainStorageName = mainStorageName;
            LastRunGuid = lastRunGuid;
        }

        /// <summary>
        /// Метод инициализации статических свойств классов для работы с ExtensibleStorage. Включая текстовую пометку от пользователя
        /// </summary>
        internal ExtensibleStorageEntity (string checkName, string mainStorageName, Guid lastRunGuid, Guid userTextGuid) : this (checkName, mainStorageName, lastRunGuid)
        {
            UserTextGuid = userTextGuid;
        }

        /// <summary>
        /// Метод инициализации статических свойств классов для работы с ExtensibleStorage. Включая текстовую пометку от пользователя и основной маркер
        /// </summary>
        internal ExtensibleStorageEntity (string checkName, string mainStorageName, Guid lastRunGuid, Guid userTextGuid, Guid markerGuid) : this (checkName, mainStorageName, lastRunGuid, userTextGuid)
        {
            MarkerGuid = markerGuid;
        }
    }
}
