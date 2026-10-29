using KPLN_Library_Forms.UI;
using KPLN_Library_Schedules.Common.Schedules;
using System;

namespace KPLN_Library_Schedules.Forms
{
    // Создаём окно его собственным типом: BAML находится в сборке KPLN_Library_Forms.
    public static class ScheduleReportForm
    {
        public static OperationReportWindow Create(ScheduleExecutionReport report)
        {
            return new OperationReportWindow("KPLN: Итоги работы со спецификациями",
                (report ?? throw new ArgumentNullException(nameof(report))).Text)
            {
                Width = 650,
                Height = 570
            };
        }
    }
}
