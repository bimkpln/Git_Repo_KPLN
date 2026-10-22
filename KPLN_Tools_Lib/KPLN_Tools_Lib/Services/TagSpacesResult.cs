using System.Collections.Generic;

namespace KPLN_Tools_Lib.Services
{
    internal sealed class TagSpacesResult
    {
        internal int Created { get; set; }
        internal int Existing { get; set; }
        internal int Skipped { get; set; }
        internal int ErrorViews { get; set; }
        internal List<string> Messages { get; } = new List<string>();
        internal List<string> Errors { get; } = new List<string>();
        internal string Summary => string.Format("Создано марок: {0}\nУже маркировано: {1}\nПропущено элементов / видов: {2}",
            Created, Existing, Skipped) + (ErrorViews > 0
                ? string.Format("\nОШИБКИ ОФОРМЛЕНИЯ: {0} план(ов). Разделите отопление и вентиляцию на разные планы.", ErrorViews)
                : string.Empty);
    }
}
