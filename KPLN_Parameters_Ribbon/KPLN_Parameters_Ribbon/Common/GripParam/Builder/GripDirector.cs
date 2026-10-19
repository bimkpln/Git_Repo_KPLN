using Autodesk.Revit.DB;
using KPLN_Library_Forms.UI;
using KPLN_Library_Forms.UI.ViewModels;
using System;

namespace KPLN_Parameters_Ribbon.Common.GripParam.Builder
{
    /// <summary>Подготавливает элементы и выполняет запись параметров захваток.</summary>
    internal class GripDirector
    {
        private readonly AbstrGripBuilder _builder;

        public GripDirector(AbstrGripBuilder builder)
        {
            _builder = builder;
        }

        public void BuildWriter()
        {
            var progress = new ProgressInfoViewModel
            {
                ProcessTitle = $"{_builder.DocMainTitle}: Параметры захваток",
                MaintStatus = "Подготавливаю элементы и геометрию",
                IsIndeterminate = true,
                UpdateIntervalMilliseconds = 125,
            };
            ProgressWindow window = null;

            try
            {
                window = ProgressWindow.ShowInBackground(progress, "KPLN: Параметры захваток");
                _builder.Prepare();
                progress.SetStage("Проверяю доступность элементов и параметров");
                _builder.Check(_builder.Doc);

                int total = _builder.ElemsOnLevel.Count + _builder.ElemsUnderLevel.Count + _builder.ElemsByHost.Count;
                Action<int, string> reportProgress = (completed, stage) => progress.Report(completed, total, stage);

                // Общая группа позволяет отменить оба этапа, сохранив отдельные транзакции для оснований.
                using (var group = new TransactionGroup(_builder.Doc, "KPLN: Параметры захваток"))
                {
                    group.Start();
                    reportProgress(0, "Поиск по геометрии");
                    using (var transaction = new Transaction(_builder.Doc, $"{_builder.DocMainTitle}: Параметры захваток_Geom"))
                    {
                        transaction.Start();
                        _builder.ExecuteGripParams_ByGeom(reportProgress);
                        progress.SetStage("Сохраняю результаты анализа геометрии");
                        if (transaction.Commit() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Не удалось сохранить параметры по геометрии.");
                    }

                    reportProgress(_builder.PbCounter, "Анализ элементов на основе");
                    using (var transaction = new Transaction(_builder.Doc, $"{_builder.DocMainTitle}: Параметры захваток_Host"))
                    {
                        transaction.Start();
                        _builder.ExecuteGripParams_ByHost(reportProgress);
                        progress.SetStage("Сохраняю результаты анализа оснований");
                        if (transaction.Commit() != TransactionStatus.Committed)
                            throw new InvalidOperationException("Не удалось сохранить параметры по основанию.");
                    }

                    progress.SetStage("Проверяю необработанные элементы");
                    _builder.CheckNotExecutedElems();
                    progress.BeginCommit();
                    if (group.Assimilate() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Заполнение параметров захваток не сохранено.");
                }

                progress.IsComplete = true;
            }
            finally
            {
                try
                {
                    window?.CloseSafely();
                }
                finally
                {
                    _builder.DisposeGeometryFilters();
                }
            }
        }
    }
}