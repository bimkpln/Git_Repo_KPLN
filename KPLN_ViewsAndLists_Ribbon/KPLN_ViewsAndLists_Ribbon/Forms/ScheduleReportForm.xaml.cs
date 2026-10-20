using KPLN_Library_Forms.UI;
using KPLN_ViewsAndLists_Ribbon.Common.Schedules;
using System;

namespace KPLN_ViewsAndLists_Ribbon.Forms
{
    // Адаптер сохраняет публичный конструктор; вся разметка и поведение находятся в Forms.
    public class ScheduleReportForm : OperationReportWindow
    {
        public ScheduleReportForm(ScheduleExecutionReport report)
            : base("KPLN: Итоги работы со спецификациями",
                (report ?? throw new ArgumentNullException(nameof(report))).Text)
        {
            Width = 650;
            Height = 570;
        }
    }
}
