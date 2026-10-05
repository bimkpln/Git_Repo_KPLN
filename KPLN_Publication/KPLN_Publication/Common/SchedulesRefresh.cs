using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Publication
{
    public static class SchedulesRefresh
    {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
        public static List<int> groupIds = new List<int>();
#else
        public static List<long> groupIds = new List<long>();
#endif
        
        public static void Start(Document doc, View sheet, bool silent = false)
        {
            List<ScheduleSheetInstance> ssis = new FilteredElementCollector(doc)
                .OfClass(typeof(ScheduleSheetInstance))
                .Cast<ScheduleSheetInstance>()
                .Where(i => i.OwnerViewId.Equals(sheet.Id))
                .Where(i => !i.IsTitleblockRevisionSchedule)
                .ToList();

            List<ScheduleSheetInstance> pinnedSchedules = new List<ScheduleSheetInstance>();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Обновление спецификаций 1");
                if (silent)
                    SetSilentFailures(t);

                foreach (ScheduleSheetInstance ssi in ssis)
                {
                    if (ssi.Pinned && (ssi.GroupId == null || ssi.GroupId == ElementId.InvalidElementId))
                    {
                        ssi.Pinned = false;
                        pinnedSchedules.Add(ssi);
                    }
                    MoveScheduleOrGroup(doc, ssi, 0.1);
                }

                if (t.Commit() != TransactionStatus.Committed && silent)
                    throw new System.InvalidOperationException("Не удалось обновить спецификации (этап 1).");
            }

            groupIds.Clear();

            using (Transaction t2 = new Transaction(doc))
            {
                t2.Start("Обновление спецификаций 2");
                if (silent)
                    SetSilentFailures(t2);

                foreach (ScheduleSheetInstance ssi in ssis)
                {
                    MoveScheduleOrGroup(doc, ssi, -0.1);
                }

                foreach (ScheduleSheetInstance ssi in pinnedSchedules)
                {
                    ssi.Pinned = true;
                }

                if (t2.Commit() != TransactionStatus.Committed && silent)
                    throw new System.InvalidOperationException("Не удалось обновить спецификации (этап 2).");
            }
        }

        private static void SetSilentFailures(Transaction transaction)
        {
            var options = transaction.GetFailureHandlingOptions();
            options.SetFailuresPreprocessor(new SilentRefreshFailures());
            options.SetClearAfterRollback(true);
            transaction.SetFailureHandlingOptions(options);
        }

        private sealed class SilentRefreshFailures : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor accessor)
            {
                return accessor.GetFailureMessages().Count == 0
                    ? FailureProcessingResult.Continue : FailureProcessingResult.ProceedWithRollBack;
            }
        }

        private static void MoveScheduleOrGroup(Document doc, ScheduleSheetInstance ssi, double distance)
        {
            if (ssi.GroupId == null || ssi.GroupId == ElementId.InvalidElementId)
            {
                ElementTransformUtils.MoveElement(doc, ssi.Id, new XYZ(distance, 0, 0));
            }
            else
            {
                Element group = doc.GetElement(ssi.GroupId);
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                if (groupIds.Contains(ssi.GroupId.IntegerValue)) return;
#else
                if (groupIds.Contains(ssi.GroupId.Value)) return;
#endif

                if (group.Pinned)
                    group.Pinned = false;

                ElementTransformUtils.MoveElement(doc, ssi.GroupId, new XYZ(distance, 0, 0));

#if Debug2020 || Revit2020 || Debug2023 || Revit2023
                groupIds.Add(ssi.GroupId.IntegerValue);
#else
                groupIds.Add(ssi.GroupId.Value);
#endif

            }
        }
    }
}
