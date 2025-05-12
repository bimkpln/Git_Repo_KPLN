using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Documents;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Класс для обработки событий Revit
    /// </summary>
    internal class RevitEventWorker : IDisposable
    {
        private static Logger _logger;
        private readonly DBRevitDialog[] _dbRevitDialogs;
        private readonly ExchangeService _exchangeEnvironment;
        private string _currentDocName;

        public RevitEventWorker(ExchangeService exchangeEnvironment, Logger logger, DBRevitDialog[] dbRevitDialogs)
        {
            _exchangeEnvironment = exchangeEnvironment;
            _logger = logger;
            _dbRevitDialogs = dbRevitDialogs;

            _exchangeEnvironment.FieldChanged += SubscribeToFieldChangedEvent;
        }

        /// <summary>
        /// Событие на попытку открытия документа
        /// </summary>
        internal void OnDocumentOpening(object sender, DocumentOpeningEventArgs args)
        {
            // Происходит, когда открытие проекта отменено
            if (args.PathName == null)
                return;

            _logger.Info($"Начало работы с файлом: {args.PathName}");
        }

        /// <summary>
        /// Событие на закрытие документа
        /// </summary>
        internal void OnDocumentClosed(object sender, DocumentClosedEventArgs args)
        {
            _logger.Info($"Конец работы с файлом: {_currentDocName}");
        }

        /// <summary>
        /// Обработка события всплывающего окна Ревит
        /// </summary>=
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs args)
        {
            _logger.Info($"Появилось окно {args.DialogId}");

            if (args.Cancellable)
            {
                args.Cancel();
                _logger.Info($"Окно {args.DialogId} успешно закрыто, благодаря стандартной возможности закрывания (Cancellable) данного окна");
            }
            else
            {
                DBRevitDialog currentDBDialog = null;
                if (string.IsNullOrEmpty(args.DialogId))
                {
                    TaskDialogShowingEventArgs taskDialogShowingEventArgs = args as TaskDialogShowingEventArgs;
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => !string.IsNullOrEmpty(rd.Message) && taskDialogShowingEventArgs.Message.Contains(rd.Message));
                }
                else
                    currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => !string.IsNullOrEmpty(rd.DialogId) && args.DialogId.Contains(rd.DialogId));

                if (currentDBDialog == null)
                {
                    _logger.Error($"Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека");
                    return;
                }

                if (Enum.TryParse(currentDBDialog.OverrideResult, out TaskDialogResult taskDialogResult))
                {
                    if (args.OverrideResult((int)taskDialogResult))
                        _logger.Info($"Окно {args.DialogId} успешно закрыто. Была применена команда {currentDBDialog.OverrideResult}");
                    else
                        _logger.Error($"Окно {args.DialogId} не удалось обработать. Была применена команда {currentDBDialog.OverrideResult}, но она не сработала!");
                }
                else
                    _logger.Error($"Не удалось привести OverrideResult '{currentDBDialog.OverrideResult}' к позиции из Autodesk.Revit.UI.TaskDialogResult. Нужна корректировка БД!");
            }
        }

        /// <summary>
        /// Обработчик ошибок. Он нужен, когда закрывание окна не работает "Error dialog has no callback"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            FailuresAccessor fa = args.GetFailuresAccessor();
            IList<FailureMessageAccessor> fmas = fa.GetFailureMessages();
            if (fmas.Count > 0)
            {
                List<FailureMessageAccessor> resolveFailures = new List<FailureMessageAccessor>();
                foreach (FailureMessageAccessor fma in fmas)
                {
                    try
                    {
                        fa.DeleteWarning(fma);
                    }
                    catch
                    {
                        //// Пытаюсь удалить элементы из ошибки, если не удалось просто удалить ошибку
                        fma.SetCurrentResolutionType(
                            fma.HasResolutionOfType(FailureResolutionType.DetachElements)
                            ? FailureResolutionType.DetachElements
                            : FailureResolutionType.DeleteElements);

                        resolveFailures.Add(fma);
                    }
                }

                if (resolveFailures.Count > 0)
                {
                    fa.ResolveFailures(resolveFailures);
                    // Убиваю окно в конце указывая коммит для обработчика
                    args.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
                }
            }
        }

        private void SubscribeToFieldChangedEvent(object sender, FieldChangedEventArgs e)
        {
            _currentDocName = e.NewValue;
        }

        public void Dispose()
        {
            _exchangeEnvironment.FieldChanged -= SubscribeToFieldChangedEvent;
            GC.SuppressFinalize(this);
        }
    }
}
