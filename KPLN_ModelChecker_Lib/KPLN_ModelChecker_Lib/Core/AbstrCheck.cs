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
    /// Статус результат процесса запуска проверки
    /// </summary>
    public enum CheckResultStatus
    {
        Failed,
        Succeeded,
        Cancelled,
        NoItemsToCheck
    }

    /// <summary>
    /// Абстрактный класс для подготовки, создания и вывода отчета пользователю
    /// </summary>
    public abstract class AbstrCheck
    {
        /// <summary>
        /// "Грязная" коллекция ошибок. Итог - в PreparedElemColl, который чиститься в основном методе "ExecuteCheck"
        /// </summary>
        private protected List<CheckerEntity> _checkerEntitiesCollHeap = new List<CheckerEntity>();

        /// <summary>
        /// Пустой конструктор для использования дженериков
        /// </summary>
        public AbstrCheck() { }

        /// <summary>
        /// Ссылка на UIApplication Revit
        /// </summary>
        public UIApplication CheckUIApp { get; private set; }

        /// <summary>
        /// Имя плагина
        /// </summary>
        public string PluginName { get; protected set; }

        /// <summary>
        /// Ссылка на ExtStorage
        /// </summary>
        public ExtensibleStorageEntity ESEntity { get; protected set; }

        /// <summary>
        /// Итоговая коллекция ошибок, полученных при проверке модели
        /// </summary>
        public CheckerEntity[] CheckerEntitiesColl { get; private protected set; }
        
        /// <summary>
        /// Ссылка на Revit-документ
        /// </summary>
        private protected Document CheckDocument { get; set; }

        /// <summary>
        /// Список подготовленных элементов, которые прошли проверку перед запуском
        /// </summary>
        private protected Element[] PreparedElemColl { get; set; }

        /// <summary>
        /// Список элементов, которые провалили проверку перед запуском
        /// </summary>
        private protected List<CheckCommandMsg> PrepareElemsErrorColl { get; set; } = new List<CheckCommandMsg>();

        /// <summary>
        /// Список элементов, которые провалили прохождение скрипта, но НЕ критичные
        /// </summary>
        private protected List<CheckCommandMsg> RunElemsWarningColl { get; set; } = new List<CheckCommandMsg>();

        /// <summary>
        /// Докрутка нужных параметров, из-за пустого конструктора для использовния дженериков
        /// </summary>
        public AbstrCheck Set_UIAppData(UIApplication uiapp, Document doc)
        {
            CheckUIApp = uiapp;
            CheckDocument = doc;

            return this;
        }

        /// <summary>
        /// Подготовить коллекцию элементов для проверки
        /// </summary>
        public abstract Element[] GetElemsToCheck();

        /// <summary>
        /// Запуск проверки
        /// </summary>
        /// <param name="elemColl">Коллеция элементов для полного анализа</param>
        /// <param name="onlyErrorType">Только сущности, с типом Error из ключевого enum по статусам проверок</param>
        /// <returns>Статус результата запуска проверки</returns>
        public CheckResultStatus ExecuteCheck(Element[] elemColl, bool onlyErrorType)
        {
            // Если нет эл-в в модели - это ОШИБКА
            if (!elemColl.Any())
            {
                CheckerEntitiesColl = new CheckerEntity[1] 
                {
                    new CheckerEntity(
                        "Отсутсвуют элементы для проверки",
                        "В модели нет подходящих для проверки элементов",
                        "Либо проверку запустили по ошибке, либо проект НЕ содержит элементы, которые должен по требованиям ВЕР") 
                };

                return CheckResultStatus.NoItemsToCheck;
            }

            try
            {
                // Проверяю элементы на критические ошибки (исключение из проверки)
                CheckRElems_SetElemErrorColl(elemColl);
                if (PrepareElemsErrorColl.Count() > 0)
                {
                    ShowElementCheckingErrorReport();
                    PreparedElemColl = elemColl.Except(PrepareElemsErrorColl.Select(e => e.MsgElement)).ToArray();
                }
                else
                    PreparedElemColl = elemColl;


                // Запускаю генерацию элементов-ошибок CheckerEntity
                CheckResultStatus setEntitiesStatus = Set_CheckerEntitiesHeap(PreparedElemColl);
                
                // Если в момент генерации были предупреждения - вывожу
                if (RunElemsWarningColl.Any())
                    ShowElemntRunWarningReport();

                // Возвращаю замечания в зависимости от потребоностей в статусе
                if (onlyErrorType)
                    CheckerEntitiesColl = _checkerEntitiesCollHeap.Where(e => e.Status == ErrorStatus.Error).ToArray();
                else
                    CheckerEntitiesColl = _checkerEntitiesCollHeap.ToArray();

                return setEntitiesStatus;
            }
            catch (Exception ex)
            {
                if (ex is CheckerException _)
                    HtmlOutput.Print($"Проверка не пройдена, работа скрипта остановлена. Исправь ошибку: {ex.Message}", MessageType.Error);
                else if (ex.InnerException != null)
                    HtmlOutput.Print($"Проверка не пройдена, работа скрипта остановлена. Передай ошибку разработчику: {ex.InnerException.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);
                else
                    HtmlOutput.Print($"Проверка не пройдена, работа скрипта остановлена. Передай ошибку разработчику: {ex.Message}. StackTrace: {ex.StackTrace}", MessageType.Error);

                return CheckResultStatus.Failed;
            }
        }

        /// <summary>
        /// Метод для подготовки и вывода отчета по ошибкам, которые были выявлены на этапе проверки элементов ПЕРЕД запуском
        /// </summary>
        public void ShowElementCheckingErrorReport()
        {
            foreach (CheckCommandMsg error in PrepareElemsErrorColl)
            {
                HtmlOutput.Print($"Элемент id {error.MsgElement.Id} не прошел проверку! Ошибка: {error.Message}", MessageType.Error);
            }
        }

        /// <summary>
        /// Вывод предупреждений пользователю об элементах, которые не прошли обработку плагином (но по сути - НЕ критичные)
        /// </summary>
        public void ShowElemntRunWarningReport()
        {
            foreach (CheckCommandMsg error in RunElemsWarningColl)
            {
                HtmlOutput.Print($"Была выявлена НЕ критическая ошибка: \n{error.Message}\n", MessageType.Warning);
            }
        }

        /// <summary>
        /// Проверка элементов перед запуском
        /// </summary>
        /// <param name="objColl">Коллеция элементов для проверки</param>
        private protected virtual void CheckRElems_SetElemErrorColl(object[] objColl) { }

        /// <summary>
        /// Запись данных в коллекцию _checkerEntitiesCollHeap, содержащая выявленные ошибки проектирования в Revit
        /// </summary>
        /// <param name="elemColl">Коллеция элементов для анализа, которые прошли проверку ПЕРЕД запуском</param>
        /// <returns>Статус результата запуска проверки</returns>
        private protected abstract CheckResultStatus Set_CheckerEntitiesHeap(Element[] elemColl);        
    }
}
