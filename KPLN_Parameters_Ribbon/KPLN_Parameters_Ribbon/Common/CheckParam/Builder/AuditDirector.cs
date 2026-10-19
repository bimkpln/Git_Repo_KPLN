using KPLN_Library_Forms.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KPLN_Parameters_Ribbon.Common.CheckParam.Builder
{
    internal class AuditDirector
    {
        /// <summary>
        /// Распорядитель - создаёт объекты, используюя builder'ы
        /// </summary>
        private readonly AbstrAuditBuilder _builder;

        public AuditDirector(AbstrAuditBuilder builder)
        {
            _builder = builder;
        }

        /// <summary>
        /// Создать builder и выполнить проверку записи параметров в элементы Ревит
        /// </summary>
        public void BuildWriter()
        {
            _builder.Prepare();

            _builder.Check();

            // Заполняю уровни с учетом секций
            using (Transaction t = new Transaction(_builder.Doc))
            {
                t.Start($"{_builder.DocMainTitle}: Проверка парамтеров");

                int max = _builder.AllElementsCount;
                var progress = new ProgressInfoViewModel
                {
                    ProcessTitle = "Проверка факта заполнения параметров",
                    MaxProgress = max,
                };
                var window = ProgressWindow.ShowInBackground(progress, "KPLN: Проверка параметров");
                try
                {
                    bool writeLevelParams = _builder.ExecuteParamsAudit(progress);
                    progress.BeginCommit();
                    if (t.Commit() != TransactionStatus.Committed)
                        throw new InvalidOperationException("Проверка параметров не завершена: транзакция отменена.");
                    progress.IsComplete = true;
                }
                finally
                {
                    window.CloseSafely();
                }
            }

            TaskDialog taskDialog = new TaskDialog("KPLN: Проверка параметров");
            taskDialog.MainInstruction = "Факт заполнения параметров данными подтвержден!";
            taskDialog.CommonButtons = TaskDialogCommonButtons.Ok;

            taskDialog.Show();
        }
    }
}
