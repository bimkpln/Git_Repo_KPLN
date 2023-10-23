using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.Common;
using KPLN_ModelChecker_User.ExecutableCommand;
using KPLN_ModelChecker_User.Forms;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Library_Forms.UI.HtmlWindow.HtmlOutput;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.ExternalCommands
{
    /// <summary>
    /// Абстрактный класс для подготовки, создания и вывода отчета пользователю
    /// </summary>
    internal abstract class AbstrCheckCommand
    {
        /// <summary>
        /// Ссылка на Revit-application
        /// </summary>
        private protected UIApplication _application;
        /// <summary>
        /// Ссылка на отчет 
        /// </summary>
        internal protected WPFReportCreator _report;
        /// <summary>
        /// Список элементов, которые провалили проверку перед запуском
        /// </summary>
        private protected IEnumerable<CheckCommandError> _errorElemCollection = new List<CheckCommandError>();
        /// <summary>
        /// Список элементов, которые прошли проверку перед запуском
        /// </summary>
        private protected Element[] _trueElemCollection;

        // Последний запуск
        private ExtensibleStorageBuilder _esBuilderRun;
        internal protected readonly string _lastRunFieldName = "Last_Run";

        // Комментарий по внесению в допустимое
        private ExtensibleStorageBuilder _esBuilderUserText;
        internal protected readonly string _userTextFieldName = "Approve_Comment";

        // Ключевой комментарий - маркер
        private protected ExtensibleStorageBuilder _esBuildergMarker;
        internal protected string _markerFieldName = "Main_Marker";
        
        /// <summary>
        /// Имя проверки
        /// </summary>
        internal static string CheckName { get; protected private set; }

        /// <summary>
        /// Имя основного Storage
        /// </summary>
        internal static string MainStorageName { get; protected private set; }

        /// <summary>
        /// GUID для Storage последнего запуска
        /// </summary>
        internal static Guid LastRunGuid { get; protected private set; }

        /// <summary>
        /// GUID для Storage комментария пользователя (для допустимых)
        /// </summary>
        internal static Guid UserTextGuid { get; protected private set; }

        /// <summary>
        /// GUID для Storage ключевого комментария
        /// </summary>
        internal static Guid MarkerGuid { get; protected private set; }

        /// <summary>
        /// Extensible Storage для последнего запуска
        /// </summary>
        private protected ExtensibleStorageBuilder ESBuilderRun
        {
            get
            {
                if (_esBuilderRun == null) _esBuilderRun = new ExtensibleStorageBuilder(LastRunGuid, _lastRunFieldName, MainStorageName);
                return _esBuilderRun;
            }
        }

        /// <summary>
        /// Extensible Storage для пользовательского комментария
        /// </summary>
        private protected ExtensibleStorageBuilder ESBuilderUserText
        {
            get
            {
                if (_esBuilderUserText == null) _esBuilderUserText = new ExtensibleStorageBuilder(LastRunGuid, _lastRunFieldName, MainStorageName);
                return _esBuilderUserText;
            }
        }

        /// <summary>
        /// Extensible Storage для ключевого комментария
        /// </summary>
        private protected ExtensibleStorageBuilder ESBuildergMarker
        {
            get
            {
                if (_esBuildergMarker == null) _esBuildergMarker = new ExtensibleStorageBuilder(MarkerGuid, _markerFieldName, MainStorageName);
                return _esBuildergMarker;
            }
        }

        /// <summary>
        /// Спец. метод для вызова данного класса из кнопки WPF: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        internal abstract Result Execute(UIApplication uiapp);

        /// <summary>
        /// Метод для инициализации и запуска всех процессов проверки
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для полного анализа</param>
        /// <returns>Коллекция WPFEntity для передачи в отчет пользовател</returns>
        internal WPFEntity[] CheckCommandRunner(Document doc, Element[] elemColl)
        {
            try
            {
                _errorElemCollection = CheckElements(doc, elemColl);
                if (_errorElemCollection.Count() > 0)
                {
                    CreateAndShowElementCheckingErrorReport(_errorElemCollection);
                    _trueElemCollection = elemColl.Except(_errorElemCollection.Select(e => e.ErrorElement)).ToArray();
                }
                else _trueElemCollection = elemColl;

                CheckAndDropExtStrApproveComment(doc, _trueElemCollection);

                return PreapareElements(doc, _trueElemCollection).ToArray();
            }
            catch (Exception ex)
            {
                if (ex is UserException userException)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = ex.Message
                    };
                    taskDialog.Show();
                }
                
                else if (ex.InnerException != null)
                    Print($"Проверка не пройдена, работа скрипта остановлена. Передай ошибку: {ex.InnerException.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);
                else
                    Print($"Проверка не пройдена, работа скрипта остановлена. Устрани ошибку: {ex.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);

                return null;
            }
        }

        /// <summary>
        /// Подготовка окна результата проверки для пользователя
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="wpfEntityColl">Коллеция WPFEntity элементов для генерации отчета</param>
        /// <param name="isMarkered">Нужно ли использовать основной маркер при создании окна?</param>
        /// <returns>Окно для вывода пользователю</returns>
        internal OutputMainForm ReportCreatorAndDemonstrator(Document doc, WPFEntity[] wpfEntityColl, bool isMarkered = false)
        {
            if (wpfEntityColl != null)
            {
                _report = CreateReport(doc, wpfEntityColl, isMarkered);

                if (_report != null)
                {
                    SetWPFEntityFiltration(_report);
                    return new OutputMainForm(_application, this.GetType().Name, _report, ESBuilderRun, ESBuilderUserText, ESBuildergMarker);
                }
            }

            return null;
        }

        /// <summary>
        /// Проверка элементов перед запуском
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для проверки</param>
        /// <returns>Коллекция CheckCommandError для элементов, которые провалили проверку</returns>
        private protected abstract IEnumerable<CheckCommandError> CheckElements(Document doc, Element[] elemColl);

        /// <summary>
        /// Анализ элементов
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для анализа, которые прошли проверку ПЕРЕД запуском</param>
        /// <returns>Коллекция WPFEntity, содержащая выявленные ошибки проектирования в Revit</returns>
        private protected abstract IEnumerable<WPFEntity> PreapareElements(Document doc, Element[] elemColl);
        
        /// <summary>
        /// Установить фильтрацию элементов в отчете
        /// </summary>
        private protected abstract void SetWPFEntityFiltration(WPFReportCreator report);

        /// <summary>
        /// Установить статус по комментарию пользователя
        /// </summary>
        /// <param name="elem">Элемент Revit</param>
        /// <param name="ifNullComment">Статус</param>
        private protected Status SetApproveStatusByUserComment(object elem, Status ifNullComment)
        {
            Status currentStatus;
            if (ESBuilderUserText.IsDataExists_Text((Element)elem))
            {
                currentStatus = Status.Approve;
            }
            else currentStatus = ifNullComment;

            return currentStatus;
        }

        /// <summary>
        /// Получить комментарий пользователя
        /// </summary>
        /// <param name="elem">Элемент Revit</param>
        /// <returns></returns>
        private protected string GetUserComment(object elem)
        {
            string approveComment = string.Empty;
            if (ESBuilderUserText.IsDataExists_Text((Element)elem)) approveComment = _esBuilderUserText.GetResMessage_Element((Element)elem).Description;

            return approveComment;
        }

        /// <summary>
        /// Метод для подготовки и вывода отчета по ошибкам, которые были выявлены на этапе проверки элементов ПЕРЕД запуском
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов, котоыре проваили проверку ПЕРЕД запуском</param>
        private void CreateAndShowElementCheckingErrorReport(IEnumerable<CheckCommandError> elemErrorColl)
        {
            foreach (CheckCommandError elem in elemErrorColl)
            {
                Print($"Элемент id {elem.ErrorElement.Id} не прошел проверку! Ошибка: {elem.ErrorMessage}", MessageType.Error);
            }
        }

        /// <summary>
        /// Проверка и очиста элементов от ложных комментариев в Extensible Storage
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для анализа</param>
        private void CheckAndDropExtStrApproveComment(Document doc, Element[] elemColl)
        {
            try
            {
                using (Transaction t = new Transaction(doc))
                {
                    t.Start($"{ModuleData.ModuleName}: Очистка ExtStr");

                    foreach (var elem in elemColl)
                    {
                        bool isDataExist = ESBuilderUserText.IsDataExists_Text(elem);
                        bool checkTrueExtStrData = ESBuilderUserText.CheckStorageDataContains_TextLog(elem, elem.Id.ToString());
                        if (isDataExist && !checkTrueExtStrData)
                            ESBuilderUserText.DropStorageData_TextLog(elem);
                    }

                    t.Commit();
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                foreach (var elem in elemColl)
                {
                    bool isDataExist = ESBuilderUserText.IsDataExists_Text(elem);
                    bool checkTrueExtStrData = ESBuilderUserText.CheckStorageDataContains_TextLog(elem, elem.Id.ToString());
                    if (isDataExist && !checkTrueExtStrData)
                        ESBuilderUserText.DropStorageData_TextLog(elem);
                }
            }
        }

        /// <summary>
        /// Метод для подготовки отчета
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="wpfEntityColl">Коллекция WPFEntity для вывода</param>
        private WPFReportCreator CreateReport(Document doc, WPFEntity[] wpfEntityColl, bool isMarkered)
        {
            if (wpfEntityColl.Any())
            {
                #region Настройка информации по логам проека
                Element piElem = doc.ProjectInformation;
                ResultMessage esMsgRun = ESBuilderRun.GetResMessage_Element(piElem);
                ResultMessage esMsgMarker = ESBuildergMarker.GetResMessage_Element(piElem);
                #endregion

                if (esMsgMarker == null) return new WPFReportCreator(wpfEntityColl, CheckName, esMsgRun.Description);
                else
                {
                    switch (esMsgMarker.CurrentStatus)
                    {
                        case MessageStatus.Ok:
                            return new WPFReportCreator(wpfEntityColl, CheckName, esMsgRun.Description, esMsgMarker.Description);

                        case MessageStatus.Error:
                            if (isMarkered)
                            {
                                TaskDialog taskDialog = new TaskDialog("[ОШИБКА]")
                                {
                                    MainInstruction = esMsgMarker.Description
                                };
                                taskDialog.Show();
                                return null;
                            }
                            else return new WPFReportCreator(wpfEntityColl, CheckName, esMsgRun.Description);
                    }
                }
            }
            else 
            { 
                Print($"[{CheckName}] Предупреждений не найдено :)", MessageType.Success);

                // Логируем последний запуск (отдельно, если все было ОК, а потом всплыли ошибки)
                ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESBuilderRun, DateTime.Now));
            } 

            return null;
        }
    }
}
