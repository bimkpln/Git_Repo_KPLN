using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>
    /// Класс для обработки событий Revit
    /// </summary>
    internal static class RevitEventWorker
    {
        /// <summary>
        /// Обработчик ошибок. Он нужен, когда закрывание окна не работает "Error dialog has no callback"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        internal static void OnFailureProcessing(object sender, FailuresProcessingEventArgs args)
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
    }
}
