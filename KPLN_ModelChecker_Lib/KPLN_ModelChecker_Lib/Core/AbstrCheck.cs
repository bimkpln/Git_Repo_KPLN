using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using KPLN_ModelChecker_Lib.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_ModelChecker_Lib.Core
{
    /// <summary>
    /// Абстрактный класс для подготовки, создания и вывода отчета пользователю
    /// </summary>
    public abstract class AbstrCheck
    {
        /// <summary>
        /// Пустой конструктор для внесения статичных данных класса
        /// </summary>
        public AbstrCheck() { }

        public string PluginName { get; protected set; }

        public ExtensibleStorageEntity ESEntity { get; protected set; }

        /// <summary>
        /// Список подготовленных элементов, которые прошли проверку перед запуском
        /// </summary>
        private protected Element[] PreparedElemColl { get; set; }

        /// <summary>
        /// Список элементов, которые провалили проверку перед запуском
        /// </summary>
        private protected IEnumerable<CheckCommandError> PrepareElemsErrorColl { get; set; } = new List<CheckCommandError>();

        /// <summary>
        /// Список элементов, которые провалили прохождение скрипта, но НЕ критичные
        /// </summary>
        private protected IEnumerable<CheckCommandError> ErrorRunColl { get; set; } = new List<CheckCommandError>();

        /// <summary>
        /// Проверка элементов перед запуском
        /// </summary>
        /// <param name="objColl">Коллеция элементов для проверки</param>
        /// <returns>Коллекция CheckCommandError для элементов, которые провалили проверку</returns>
        private protected abstract IEnumerable<CheckCommandError> CheckRElems(object[] objColl);

        /// <summary>
        /// Получить коллекцию CheckerEntity
        /// </summary>
        /// <param name="doc">Revit-проект</param>
        /// <param name="elemColl">Коллеция элементов для анализа, которые прошли проверку ПЕРЕД запуском</param>
        /// <returns>Коллекция WPFEntity, содержащая выявленные ошибки проектирования в Revit</returns>
        private protected abstract IEnumerable<CheckerEntity> GetCheckerEntities(Document doc, Element[] elemColl);

        /// <summary>
        /// Подготовить коллекцию элементов для проверки
        /// </summary>
        public abstract Element[] GetElemsToCheck(Document doc);

        /// <summary>
        /// Запуск проверки
        /// </summary>
        /// <param name="doc">Revit-документ</param>
        /// <param name="elemColl">Коллеция элементов для полного анализа</param>
        /// <returns>Коллекция CheckerEntity для передачи в отчет пользовател</returns>
        public CheckerEntity[] ExecuteCheck(Document doc, Element[] elemColl)
        {
            if (!elemColl.Any()) return null;

            try
            {
                PrepareElemsErrorColl = CheckRElems(elemColl);
                if (PrepareElemsErrorColl.Count() > 0)
                    PreparedElemColl = elemColl.Except(PrepareElemsErrorColl.Select(e => e.ErrorElement)).ToArray();
                else
                    PreparedElemColl = elemColl;

                CheckerEntity[] result = GetCheckerEntities(doc, PreparedElemColl).ToArray();
                return result;
            }
            catch (Exception ex)
            {
                if (ex is CheckerException _)
                {
                    TaskDialog taskDialog = new TaskDialog("ОШИБКА: Выполни инструкцию")
                    {
                        MainContent = ex.Message
                    };
                    taskDialog.Show();
                }

                else if (ex.InnerException != null)
                    HtmlOutput.Print($"Проверка не пройдена, работа скрипта остановлена. Передай ошибку: {ex.InnerException.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);
                else
                    HtmlOutput.Print($"Проверка не пройдена, работа скрипта остановлена. Устрани ошибку: {ex.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);

                return null;
            }
        }

        /// <summary>
        /// Метод для подготовки и вывода отчета по ошибкам, которые были выявлены на этапе проверки элементов ПЕРЕД запуском
        /// </summary>
        public void ShowElementCheckingErrorReport()
        {
            foreach (CheckCommandError error in PrepareElemsErrorColl)
            {
                HtmlOutput.Print($"Элемент id {error.ErrorElement.Id} не прошел проверку! Ошибка: {error.ErrorMessage}", MessageType.Error);
            }
        }

        /// <summary>
        /// Вывод предупреждений пользователю об элементах, которые не прошли обработку плагином (но по сути - НЕ критичные)
        /// </summary>
        public void ShowErrorRunColl()
        {
            foreach (CheckCommandError error in ErrorRunColl)
            {
                HtmlOutput.Print($"Была выявлена НЕ критическая ошибка: \n{error.ErrorMessage}\n", MessageType.Warning);
            }
        }
    }
}
