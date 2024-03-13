using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using NLog;
using System;
using System.Linq;

namespace KPLN_BIMTools_Ribbon.Common
{
    /// <summary>
    /// Класс для обработки событий Revit
    /// </summary>
    internal class RevitEventWorker : IDisposable
    {
        private static Logger _logger;
        private readonly DBRevitDialog[] _dbRevitDialogs;
        private readonly ExchangeEnvironment _exchangeEnvironment;
        private string _currentDocName;

        public RevitEventWorker(ExchangeEnvironment exchangeEnvironment, Logger logger, DBRevitDialog[] dbRevitDialogs)
        {
            _exchangeEnvironment = exchangeEnvironment;
            _logger = logger;
            _dbRevitDialogs = dbRevitDialogs;

            _exchangeEnvironment.FieldChanged += SubscribeToFieldChangedEvent;
        }

        /// <summary>
        /// Событие на открытие документа
        /// </summary>
        internal void OnDocumentOpened(object sender, DocumentOpenedEventArgs args)
        {
            _logger.Info($"Начало работы с файлом: {args.Document.PathName}");
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
                DBRevitDialog currentDBDialog = _dbRevitDialogs.FirstOrDefault(rd => args.DialogId.Contains(rd.DialogId));
                if (currentDBDialog != null)
                {
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
                else
                {
                    _logger.Error($"Окно {args.DialogId} не удалось обработать. Необходим контроль со стороны человека");
                }
            }
        }

        private void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
        {
            // Пока не понятно, нужна ли реализация
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
