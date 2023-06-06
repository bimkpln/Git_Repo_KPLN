using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_ExtensibleStorage;
using KPLN_ModelChecker_User.WPFItems;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Loader.Output.Output;
using static KPLN_ModelChecker_User.Common.Collections;

namespace KPLN_ModelChecker_User.Forms
{
    /// <summary>
    /// Абстрактный класс для создания и вывода подготовленного отчета пользователю
    /// </summary>
    internal abstract class AbstrUserOutput
    {
        /// <summary>
        /// Имя проверки
        /// </summary>
        private protected string _name;
        /// <summary>
        /// Ссылка на Revit-application
        /// </summary>
        private protected UIApplication _application;
        /// <summary>
        /// Ссылка на отчет 
        /// </summary>
        private protected WPFReportCreator _report;

        // Последний запуск
        private ExtensibleStorageBuilder _esBuilderRun;
        private protected Guid _lastRunGuid;
        private protected string _lastRunFieldName;
        private protected string _lastRunStorageName;

        // Комментарий по внесению в допустимое
        private ExtensibleStorageBuilder _esBuilderUserText;
        private protected Guid _userTextGuid;
        private protected string _userTextFieldName;
        private protected string _userTextStorageName;

        // Ключевой комментарий - маркер
        private ExtensibleStorageBuilder _esBuildergMarker;
        private protected Guid _markerGuid;
        private protected string _markerFieldName;
        private protected string _markerStorageName;

        /// <summary>
        /// Extensible Storage для последнего запуска
        /// </summary>
        private protected ExtensibleStorageBuilder ESBuilderRun
        {
            get
            {
                if (_esBuilderRun == null) _esBuilderRun = new ExtensibleStorageBuilder(_lastRunGuid, _lastRunFieldName, _lastRunStorageName);
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
                if (_esBuilderUserText == null) _esBuilderUserText = new ExtensibleStorageBuilder(_lastRunGuid, _lastRunFieldName, _lastRunStorageName);
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
                if (_esBuildergMarker == null) _esBuildergMarker = new ExtensibleStorageBuilder(_markerGuid, _markerFieldName, _markerStorageName);
                return _esBuildergMarker;
            }
        }

        /// <summary>
        /// Спец. метод для вызова данного класса из кнопки WPF: https://thebuildingcoder.typepad.com/blog/2016/11/using-other-events-to-execute-add-in-code.html#:~:text=anything%20with%20documents.-,Here%20is%20an%20example%20code%20snippet%3A,-public%C2%A0class
        /// </summary>
        internal abstract Result Execute(UIApplication uiapp);

        /// <summary>
        /// Проверка и подготовка (при необходиомости) элементов перед запуском
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для проверки</param>
        internal abstract void CheckElements(Document doc, IEnumerable<Element> elemColl);

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
        /// Метод для подготовки и проверки необходимости отчета
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="wpfEntityColl">Коллекция WPFEntity для вывода</param>
        /// /// <returns>True, если необходимость в отчете есть (он НЕ пустой)</returns>
        private protected bool CreateAndCheckReport(Document doc, IEnumerable<WPFEntity> wpfEntityColl)
        {
            if (wpfEntityColl.Any())
            {
                ProjectInfo pi = doc.ProjectInformation;
                Element piElem = pi as Element;
                ResultMessage esMsgRun = ESBuilderRun.GetResMessage_Element(piElem);

                _report = new WPFReportCreator(
                    wpfEntityColl,
                    _name,
                    esMsgRun.Description);
                return true;
            }
            else
            {
                Print($"[{_name}] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                return false;
            }
        }

        /// <summary>
        /// Метод для подготовки и проверки необходимости отчета
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="wpfEntityColl">Коллекция WPFEntity для вывода</param>
        /// <param name="esMsgMarker">Extensible Storage: данные по ключевому логу</param>
        /// <returns>True, если необходимость в отчете есть (он НЕ пустой)</returns>
        private protected bool CreateAndCheckReport(Document doc, List<WPFEntity> wpfEntityColl, ResultMessage esMsgMarker)
        {
            if (wpfEntityColl.Any())
            {
                ProjectInfo pi = doc.ProjectInformation;
                Element piElem = pi as Element;
                ResultMessage esMsgRun = ESBuilderRun.GetResMessage_Element(piElem);

                _report = new WPFReportCreator(
                    wpfEntityColl,
                    _name,
                    esMsgRun.Description,
                    esMsgMarker.Description);
                return true;
            }
            else
            {
                Print($"[{_name}] Предупреждений не найдено!", KPLN_Loader.Preferences.MessageType.Success);
                return false;
            }
        }
    }
}
