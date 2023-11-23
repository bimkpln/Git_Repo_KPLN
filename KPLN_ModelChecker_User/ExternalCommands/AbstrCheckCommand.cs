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
    internal abstract class AbstrCheckCommand<T>
    {
        /// <summary>
        /// Ссылка на Revit-UIApplication
        /// </summary>
        private protected UIApplication _uiApp;
        /// <summary>
        /// Ссылка на отчет 
        /// </summary>
        internal protected WPFReportCreator _report;
        /// <summary>
        /// Список элементов, которые провалили проверку перед запуском
        /// </summary>
        private protected IEnumerable<CheckCommandError> _errorCheckElemsColl = new List<CheckCommandError>();
        /// <summary>
        /// Список элементов, которые провалили прохождение скрипта, но НЕ критичные
        /// </summary>
        private protected IEnumerable<CheckCommandError> _errorRunColl = new List<CheckCommandError>();
        /// <summary>
        /// Список элементов, которые прошли проверку перед запуском
        /// </summary>
        private protected Element[] _trueElemCollection;

        /// <summary>
        /// Конструктор для классов, наследуемых от AbstrCheckCommand. Если его не переопределить в наследнике - IExternalCommand не справиться с запуском (ему нужен конструтор по умолчанию)
        /// </summary>
        public AbstrCheckCommand()
        {
        }

        /// <summary>
        /// Конструктор для класса Module. Он инициализирует основные переменные для работы с ExtensibleStorage
        /// </summary>
        internal AbstrCheckCommand(ExtensibleStorageEntity esEntity)
        {
            ESEntity = esEntity;
        }

        internal static ExtensibleStorageEntity ESEntity { get; private protected set; }

        /// <summary>
        /// Спец. метод для вызова данного класса из кнопки WPF: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        public abstract Result ExecuteByUIApp(UIApplication uiapp);

        /// <summary>
        /// Метод для инициализации и запуска всех процессов проверки
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для полного анализа</param>
        /// <returns>Коллекция WPFEntity для передачи в отчет пользовател</returns>
        internal WPFEntity[] CheckCommandRunner(Document doc, Element[] elemColl)
        {
            
            if (ESEntity.CheckName == null || string.IsNullOrEmpty(ESEntity.CheckName))
                throw new ArgumentNullException("Ты забыл инициировать основные поля. Это вынесено в класс Module, чтобы обработать плагином по выводу информации по запускам скриптов");
            
            try
            {
                _errorCheckElemsColl = CheckElements(doc, elemColl);
                if (_errorCheckElemsColl.Count() > 0)
                {
                    CreateAndShowElementCheckingErrorReport(_errorCheckElemsColl);
                    _trueElemCollection = elemColl.Except(_errorCheckElemsColl.Select(e => e.ErrorElement)).ToArray();
                }
                else 
                    _trueElemCollection = elemColl;

                CheckAndDropExtStrApproveComment(doc, _trueElemCollection);

                WPFEntity[] result = PreapareElements(doc, _trueElemCollection).ToArray();
                ShowErrorRunColl();
                return result;
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
        /// Вывод предупреждений пользователю об элементах, которые не прошли обработку плагином (но по сути - НЕ критичные)
        /// </summary>
        internal void ShowErrorRunColl()
        {
            if (_errorRunColl.Any())
            {
                foreach (CheckCommandError error in _errorRunColl)
                {
                    Print($"Была выявлена НЕ критическая ошибка: \n{error.ErrorMessage}\n", MessageType.Warning);
                }
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
                    return new OutputMainForm(_uiApp, this.GetType().Name, _report, ESEntity.ESBuilderRun, ESEntity.ESBuilderUserText, ESEntity.ESBuildergMarker);
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
        private protected abstract IEnumerable<CheckCommandError> CheckElements(Document doc, object[] objColl);

        /// <summary>
        /// Анализ элементов Revit и подготовка элементов WPFEntity для отчета
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
        /// <param name="obj">Объект, который должен представлять элемент Revit</param>
        /// <param name="ifNullComment">Статус</param>
        private protected Status SetApproveStatusByUserComment(object obj, Status ifNullComment)
        {
            Status currentStatus;
            if (obj is Element elem)
            {
                if (ESEntity.ESBuilderUserText.IsDataExists_Text(elem))
                {
                    currentStatus = Status.Approve;
                }
                else currentStatus = ifNullComment;
            }
            else
                throw new Exception($"{obj} - не Element Revit");
            

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
            if (ESEntity.ESBuilderUserText.IsDataExists_Text((Element)elem)) approveComment = ESEntity.ESBuilderUserText.GetResMessage_Element((Element)elem).Description;

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
                        bool isDataExist = ESEntity.ESBuilderUserText.IsDataExists_Text(elem);
                        bool checkTrueExtStrData = ESEntity.ESBuilderUserText.CheckStorageDataContains_TextLog(elem, elem.Id.ToString());
                        if (isDataExist && !checkTrueExtStrData)
                            ESEntity.ESBuilderUserText.DropStorageData_TextLog(elem);
                    }

                    t.Commit();
                }
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                foreach (var elem in elemColl)
                {
                    bool isDataExist = ESEntity.ESBuilderUserText.IsDataExists_Text(elem);
                    bool checkTrueExtStrData = ESEntity.ESBuilderUserText.CheckStorageDataContains_TextLog(elem, elem.Id.ToString());
                    if (isDataExist && !checkTrueExtStrData)
                        ESEntity.ESBuilderUserText.DropStorageData_TextLog(elem);
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
                ResultMessage esMsgRun = ESEntity.ESBuilderRun.GetResMessage_Element(piElem);
                ResultMessage esMsgMarker = ESEntity.ESBuildergMarker.GetResMessage_Element(piElem);
                #endregion

                if (ESEntity.ESBuildergMarker.Guid.Equals(Guid.Empty)) 
                    return new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description);
                else
                {
                    switch (esMsgMarker.CurrentStatus)
                    {
                        case MessageStatus.Ok:
                            return new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description, esMsgMarker.Description);

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
                            else return new WPFReportCreator(wpfEntityColl, ESEntity.CheckName, esMsgRun.Description);
                    }
                }
            }
            else 
            { 
                Print($"[{ESEntity.CheckName}] Предупреждений не найдено :)", MessageType.Success);

                // Логируем последний запуск (отдельно, если все было ОК, а потом всплыли ошибки)
                ModuleData.CommandQueue.Enqueue(new CommandWPFEntity_SetTimeRunLog(ESEntity.ESBuilderRun, DateTime.Now));
            } 

            return null;
        }
    }
}
